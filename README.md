# 1D and 2D Barcode generation VBA macro

VBA macro able to generate the 1D [Code128][code128], [EAN][ean],
[Interleaved 2 of 5][i2o5] and [Code39][code39] barcodes, and the 2D
[Data Matrix][data-matrix] and [QR Code][qr-code] barcodes.

The macro is able to generate the barcodes as graphical objects or as
text which will be shown as the proper barcode when the companion
[BarsAndSpaces.ttf](BarsAndSpaces.ttf) font is applied to the text.

## Usage

The syntax of the macro is as follows:

```basic
ENCODEBARCODE(sheetIndex; cellAddress; code; codeType; outputType; params; quietZoneWidth)
```

- `sheetIndex`: index of the sheet in which the graphical object will
  be placed, in case of selecting that output type.
- `cellAddress`: cell address in which the graphical object will be
  placed, in case of selecting that output type.
- `code`: content to encode into the barcode.
- `codeType`: barcode type to use. Possible values:
  - `0`: [Code128](#code128) **(default)**.
  - `1`: [EAN](#ean).
  - `2`: [Interleaved 2 of 5](#interleaved-2-of-5).
  - `3`: [Code39](#code39).
  - `50`: [Data Matrix](#data-matrix).
  - `51`: [QR Code](#qr-code).
- `outputType`: whether to generate a graphical object or just the
  text for which to apply the companion font. Possible values:
  - `0`: use text.
  - `1`: use a graphic object **(default)**.
- `params`: codeType specific parameters. See below.
- `quietZoneWidth`: width of the 1D codes Quiet Zone. Default value:
  `2`.

Example:

```basic
ENCODEBARCODE(CELL("SHEET"); CELL("ADDRESS"); A2; 1; 1; 0; 2)
```

### Code128

`params` possible values:

- `0`: determine the [code set][code-set] to use automatically
  **(default)**.
- `1`: start with code set A.
- `2`: start with code set B.
- `3`: start with code set C.

[code128]: https://en.wikipedia.org/wiki/Code_128
[code-set]: https://en.wikipedia.org/wiki/Code_128#Subtypes

### EAN

`params` possible values:

- `0`: automatically determine the subtype to use **(default)**.
- `1`: use the EAN-13 subtype.
- `2`: use the [EAN-8][ean-8] subtype.
- `3`: use the [UPC-A][upc-a] subtype.
- `4`: use the [UPC-E][upc-e] subtype.
- To add the [check digit][check-digit], increase the chosen subtype
  by `8`. For example, a value of `9` means to use the EAN-13 subtype
  and add the check digit.

[ean]: https://en.wikipedia.org/wiki/International_Article_Number
[ean-8]: https://en.wikipedia.org/wiki/EAN-8
[upc-a]: https://en.wikipedia.org/wiki/Universal_Product_Code
[upc-e]: https://en.wikipedia.org/wiki/Universal_Product_Code#UPC-E
[check-digit]: https://en.wikipedia.org/wiki/International_Article_Number#Check_digit

### Interleaved 2 of 5

No additional parameters.

[i2o5]: https://en.wikipedia.org/wiki/Interleaved_2_of_5

### Code39

`params` possible values:

- `0`: automatically use the [Full ASCII Code 39][full-ascii-code-39]
  mode.
- `1`: always use the Full ASCII Code 39 mode **(default)**.
- `2`: don't use the Full ASCII Code 39 mode.
- To add the [Code 39 mod 43][code-39-mo-43] check digit, increase the
  chosen mode by `4`. For example, a value of `5` means to always use
  the Full ASCII Code 39 mode and add the check digit.

[code39]: https://en.wikipedia.org/wiki/Code_39
[full-ascii-code-39]: https://en.wikipedia.org/wiki/Code_39#Full_ASCII_Code_39
[code-39-mo-43]: https://en.wikipedia.org/wiki/Code_39#Code_39_mod_43

### Data Matrix

`params` possible values:

- `0`: autodetect the encoding mode ([ASCII][ascii],
  [C40 and Text][text-modes]) **(default)**.
- `1`: force ASCII encoding mode.

[data-matrix]: https://en.wikipedia.org/wiki/Data_Matrix
[ascii]: https://en.wikipedia.org/wiki/Data_Matrix#Encoding
[text-modes]: https://en.wikipedia.org/wiki/Data_Matrix#Text_modes

### QR Code

`params` possible values:

- `0`: [Error Correction (EC) level][error-correction] *M*
  **(default)**.
- `1`: EC level *L*.
- `2`: EC level *Q*.
- `3`: EC level *H*.

[qr-code]: https://en.wikipedia.org/wiki/QR_code
[error-correction]: https://en.wikipedia.org/wiki/QR_code#Error_correction

## Similar projects

- [OPENBARCODES project](http://grandzebu.net/informatique/codbar-en/codbar.htm)
