# pptx2imgs

Create a PPTX with one full-slide PNG per slide. Slide size, order, and hidden status are preserved.

Install Python 3.10+ and LibreOffice Impress. In a Python virtual environment, run:

```sh
python -m pip install -r requirements.txt
python pptx2imgs.py "deck.pptx"
```

The output is `deck-images.pptx`. Slides use lossless PNGs and render at 300 DPI or higher to retain detail from embedded images. PowerPoint's automatic image compression is disabled in the output.

Use `-o result.pptx` to choose the output file and `--dpi 600` to set a fixed render resolution.

On NixOS, run the converter inside `nix-shell -p libreoffice 'python3.withPackages (ps: [ ps.pymupdf ps.python-pptx ])'`.

LibreOffice uses the fonts available on the conversion computer.
