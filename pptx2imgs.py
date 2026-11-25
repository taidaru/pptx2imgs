import argparse
import os
import posixpath
import shutil
import subprocess
import tempfile
from io import BytesIO
from math import ceil, hypot
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile


DEFAULT_DPI = 300
PDF_FILTER = (
    'pdf:impress_pdf_Export:{"ExportHiddenSlides":{"type":"boolean","value":"true"},'
    '"UseLosslessCompression":{"type":"boolean","value":"true"},'
    '"ReduceImageResolution":{"type":"boolean","value":"false"}}'
)


class ConversionError(Exception):
    pass


def find_libreoffice():
    for name in ("soffice.com", "libreoffice", "soffice"):
        executable = shutil.which(name)
        if executable:
            return executable

    locations = [Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")]
    if os.name == "nt":
        for variable in ("ProgramFiles", "ProgramFiles(x86)"):
            if program_files := os.environ.get(variable):
                locations.append(Path(program_files) / "LibreOffice" / "program" / "soffice.com")

    for location in locations:
        if location.is_file():
            return str(location)
    raise ConversionError("LibreOffice Impress is required. Install it or add soffice to PATH.")


def export_pdf(source, workspace, libreoffice):
    result = subprocess.run(
        [
            libreoffice,
            f"-env:UserInstallation={(workspace / 'profile').as_uri()}",
            "--headless",
            "--infilter=Impress MS PowerPoint 2007 XML",
            "--convert-to",
            PDF_FILTER,
            "--outdir",
            str(workspace),
            str(source),
        ],
        capture_output=True,
        text=True,
        errors="replace",
    )
    pdf = workspace / source.with_suffix(".pdf").name
    if result.returncode or not pdf.is_file() or not pdf.stat().st_size:
        detail = (result.stderr or result.stdout).strip()
        raise ConversionError(f"LibreOffice could not export '{source.name}'.\n{detail}".rstrip())
    return pdf


def read_slide_properties(source):
    try:
        with ZipFile(source) as package:
            package_rels = ElementTree.fromstring(package.read("_rels/.rels"))
            main_target = next(
                rel.attrib["Target"]
                for rel in package_rels
                if rel.get("Type", "").endswith("/officeDocument")
            )
            main_part = PurePosixPath(posixpath.normpath(main_target).lstrip("/"))
            presentation = ElementTree.fromstring(package.read(str(main_part)))
            size = presentation.find("{*}sldSz")
            dimensions = None if size is None else (int(size.attrib["cx"]), int(size.attrib["cy"]))
            slide_ids = presentation.findall("{*}sldIdLst/{*}sldId")
            if not slide_ids:
                raise ConversionError("The presentation contains no slides.")

            rels_part = main_part.parent / "_rels" / f"{main_part.name}.rels"
            slide_rels = ElementTree.fromstring(package.read(str(rels_part)))
            targets = {rel.attrib["Id"]: rel.attrib["Target"] for rel in slide_rels}

            hidden_slides = []
            for slide in slide_ids:
                rel_id = next(value for name, value in slide.attrib.items() if name.endswith("}id"))
                slide_part = posixpath.normpath(posixpath.join(str(main_part.parent), targets[rel_id]))
                with package.open(slide_part.lstrip("/")) as stream:
                    _, root = next(ElementTree.iterparse(stream, events=("start",)))
                    hidden_slides.append(root.get("show") in ("0", "false"))
    except (BadZipFile, ElementTree.ParseError, KeyError, StopIteration, ValueError) as error:
        raise ConversionError(f"Could not read the slides in '{source.name}'.") from error

    return dimensions, hidden_slides


def save_presentation(presentation, output):
    output.parent.mkdir(parents=True, exist_ok=True)
    stream = output.open("xb")
    try:
        with stream:
            presentation.save(stream)
    except BaseException:
        output.unlink(missing_ok=True)
        raise


def recommended_dpi(page):
    dpi = DEFAULT_DPI
    for image in page.get_image_info():
        a, b, c, d, _, _ = image["transform"]
        width = hypot(a, b)
        height = hypot(c, d)
        if width > 0 and height > 0:
            dpi = max(dpi, ceil(72 * image["width"] / width), ceil(72 * image["height"] / height))
    return dpi


def convert(source, output, dpi):
    try:
        import pymupdf
        from PIL import Image
        from pptx import Presentation
        from pptx.util import Pt
    except ImportError as error:
        raise ConversionError(
            "Could not load a dependency. Run 'python -m pip install -r requirements.txt'.\n"
            f"{error}"
        ) from error

    if output.exists() or output.is_symlink():
        raise FileExistsError(f"Output path already exists: {output}")

    dimensions, hidden_slides = read_slide_properties(source)
    count = len(hidden_slides)
    presentation = Presentation()
    presentation.element.set("autoCompressPictures", "0")
    blank = presentation.slide_layouts.get_by_name("Blank")

    libreoffice = find_libreoffice()
    with tempfile.TemporaryDirectory(prefix="pptx2imgs-") as temp:
        pdf = export_pdf(source, Path(temp).resolve(), libreoffice)
        with pymupdf.open(pdf) as document:
            if len(document) != count:
                raise ConversionError(f"LibreOffice rendered {len(document)} of {count} slides.")

            if dimensions is None:
                page_size = document[0].rect
                dimensions = Pt(page_size.width), Pt(page_size.height)
            width, height = dimensions
            presentation.slide_width = width
            presentation.slide_height = height

            for page, hidden in zip(document, hidden_slides):
                slide_dpi = recommended_dpi(page) if dpi is None else dpi
                bounds = (page.rect * (slide_dpi / 72)).irect
                max_pixels = Image.MAX_IMAGE_PIXELS
                if max_pixels is not None and bounds.width * bounds.height > 2 * max_pixels:
                    raise ConversionError(
                        f"Slide {page.number + 1} at {slide_dpi} DPI exceeds the supported PNG size. "
                        "Use --dpi to choose a lower resolution."
                    )
                slide = presentation.slides.add_slide(blank)
                if hidden:
                    slide.element.set("show", "0")
                png = BytesIO(page.get_pixmap(dpi=slide_dpi).tobytes("png"))
                slide.shapes.add_picture(png, 0, 0, width, height)

    save_presentation(presentation, output)
    return count


def main():
    parser = argparse.ArgumentParser(description="Create a PPTX with one full-slide PNG per slide.")
    parser.add_argument("presentation", type=Path, help="Path to a .pptx file")
    parser.add_argument("-o", "--output", type=Path, help="Path to a new .pptx file")
    parser.add_argument("--dpi", type=int, help=f"Fixed render resolution (default: adaptive, at least {DEFAULT_DPI})")
    args = parser.parse_args()

    source = args.presentation.expanduser().absolute()
    if source.suffix.lower() != ".pptx" or not source.is_file():
        parser.error(f"Expected an existing .pptx file: {source}")
    if args.dpi is not None and args.dpi <= 0:
        parser.error("--dpi must be a positive integer")

    output = source.with_name(f"{source.stem}-images.pptx")
    if args.output is not None:
        output = args.output.expanduser().absolute()
    if output.suffix.lower() != ".pptx":
        parser.error("--output must be a .pptx file")
    try:
        count = convert(source, output, args.dpi)
    except (ConversionError, OSError, RuntimeError, ValueError) as error:
        parser.exit(1, f"Error: {error}\n")

    print(f"Saved {count} image slide(s) to {output}")


if __name__ == "__main__":
    main()
