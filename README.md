# pptx2imgs

Create a PPTX with one full-slide PNG per slide. Slide size, order, and hidden status are preserved.

Install Python 3.10+ and LibreOffice Impress. In a Python virtual environment, run:

```sh
python -m pip install -r requirements.txt
python pptx2imgs.py "deck.pptx"
```

The output is `deck-images.pptx`. Use `-o result.pptx --dpi 200` to choose a new file and change the default resolution of 150 DPI.

On NixOS, run the converter inside `nix-shell -p libreoffice 'python3.withPackages (ps: [ ps.pymupdf ps.python-pptx ])'`.

LibreOffice uses the fonts available on the conversion computer.
