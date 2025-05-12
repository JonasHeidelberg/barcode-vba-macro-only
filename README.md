# barcode-vba-macro-only
Barcode generator as pure VBA macro

A bug which sometimes caused incorrect output ("stuttering") has been fixed here, but only tested for MSOffice. The original 2013 version by Jiri Gabriel contained files for LibreOffice and MSOffice

New in 2025: Small & simple xlsm file for MS Office, only for creating QR codes (other barcodes thrown out). Features: 
- Uses "Update" button instead of formula to reduce CPU load
- Converts output from Shapes to SVG, to enable easier export e.g. to PowerPoint
- For the input cell, prioritizes the hyperlink over the displayed cell content
