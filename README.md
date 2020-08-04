# DocX Builder

This is a package for making DocX files with OOXML for Flutter.
You can make .docx files with html+css by simply changing the .html file extension to .docx.
This might not be an option however in some cases, for example:

- you do not have or want to use html+css
- you need to be sure the document looks uniform and exactly the way as described (word processors can show unpredictable and different results for the same html file)
- you need some features html+css do not offer, e.g. add headers/footers/comments/certain dynamic information like page numbers.

This package can then be used.

## Getting started

In pubspec.yaml, get this package:

```yaml
docx_builder:
  git:
    url: https://github.com/VergillR/docx_builder
```

Then import it:

```dart
import 'package:path_provider/path_provider.dart';
import 'package:docx_builder/docx_builder.dart' as dx;

// initialize DocXBuilder by giving it a cache directory to store temporary files

final tempDir = await getTemporaryDirectory();

final docx = dx.DocXBuilder(tempDir);
```

If you are using a background color, headers, footers or global styles, you can set them up now. For example:

```dart
docx.setDocumentBackgroundColor('AABBFF');

docx.globalTextStyle = dx.TextStyle(
    fontSize: 20,
    fontName: 'Arial',
    tabs: <dx.DocxTab>[
      dx.DocxTab(
        leader: dx.TabLeader.dot,
        position: 2100,
      ),
    ]);

docx.setFooter(
  dx.FooterType.oddPage,
  ['Page number: '],
  [
    dx.TextStyle(
      fontSize: 12,
      bold: true,
      italic: true,
      textAlignment: dx.TextAlignment.center,
      fontColor: 'DFDFDF',
      highlightColor: dx.HighlightColor.darkBlue,
    ),
  ],
  doNotUseGlobalTextStyle: true,
  complexFields: [ dx.ComplexField(instructions: ' PAGE ')],
);

docx.attachHeadersAndFooters();
```

Adding text is done with either addText (all text use the same style) or addMixedText (text can use different styles).

```dart
docx.addText(
  'Document header',
  textStyle: dx.TextStyle(
    fontSize: 24.0,
    fontColor: 'DF2365',
    bold: true,
  ),
);

docx.addMixedText(
  <String>[
    "This is normal text with ",
    "a link to the Flutter website ",
    "as an example.",
  ],
  // this is a list of text styles
  // null means that global style should be used
  <dx.TextStyle>[
    null,
    dx.TextStyle(
      hyperlinkTo: 'https://www.flutter.dev',
    ),
    null,
  ],
);
```

Images can be added with addImage, addImageWithCaption, addImages or addAnchorImage. For example:

```dart
docx.addImage(
  File('${appDir.path}/test.jpg'),
  2600000,
  3000000,
  alternativeTextForImage: 'The user can click on this image to go to the Flutter website',
  hyperlinkTo: 'https://www.flutter.dev',
);
```

Tables are complex structures in OOXML. The steps are:

1. Create table cells and put text or images inside them with insertTextInTableCell and insertImageInTableCell.
2. Create table rows and put the table cells inside the rows.
3. Create a table and put the table rows inside the table.
4. Attach the table by calling attachTable.

```dart
final dx.TableCell tc1 = dx.TableCell();
docx.insertTextInTableCell(
    tableCell: tc1,
    text: 'This is tc1 cell.');
final dx.TableCell tc2 = dx.TableCell();
docx.insertTextInTableCell(
    textStyle: dx.TextStyle(
      numberingList: dx.NumberingList.bullet,
      numberLevelInList: 1,
    ),
    tableCell: tc2,
    text: 'This is bulleted text in tc2.');
final dx.TableCell tc3 = dx.TableCell();
docx.insertTextInTableCell(
    tableCell: tc3,
    text: 'This is tc3 cell.');
final dx.TableCell tc4 = dx.TableCell();
docx.insertMixedTextInTableCell(
    tableCell: tc4,
    text: [
      'This is mixed text in ',
      'tc4 cell.'
    ],
    textStyles: [
      dx.TextStyle(bold: true),
      dx.TextStyle(
        bold: true,
        italic: true,
        highlightColor: dx.HighlightColor.darkGreen,
      ),
    ],
    doNotUseGlobalTextStyle: true);
final dx.TableCell tc5 = dx.TableCell(
    tableCellProperties:
        dx.TableCellProperties(gridSpan: 2));
docx.insertImagesInTableCell(
  imageFiles: [
    File('${appDir.path}/test1.jpg'),
    File('${appDir.path}/test2.jpg'),
  ],
  widthEMU: 2600000,
  heightEMU: 3000000,
  tableCell: tc5,
);

final dx.TableRow row1 =
    dx.TableRow(tableCells: [tc1, tc1, tc2]);
final dx.TableRow row2 =
    dx.TableRow(tableCells: [tc3, tc3, tc3]);
final dx.TableRow row3 = dx.TableRow(tableCells: [
  tc4,
  tc5,
]);

final dx.Table table = dx.Table(
  gridColumnWidths: [2000, 2500, 3300],
  tableRows: [row1, row2, row3],
  tableProperties: dx.TableProperties(
    tableBorders: [
      dx.TableBorder(
        borderSide: dx.TableBorderSide.left,
        color: '3434DF',
        pbrStyle: dx.ParagraphBorderStyle.double,
        width: 4,
      ),
      dx.TableBorder(
        borderSide: dx.TableBorderSide.right,
        color: '3434DF',
        pbrStyle: dx.ParagraphBorderStyle.double,
        width: 4,
      ),
      dx.TableBorder(
        borderSide: dx.TableBorderSide.top,
        color: '3434DF',
        pbrStyle: dx.ParagraphBorderStyle.double,
        width: 4,
      ),
      dx.TableBorder(
        borderSide: dx.TableBorderSide.bottom,
        color: '3434DF',
        pbrStyle: dx.ParagraphBorderStyle.double,
        width: 4,
      ),
      dx.TableBorder(
        borderSide: dx.TableBorderSide.insideH,
        color: '3434DF',
        pbrStyle: dx.ParagraphBorderStyle.single,
        width: 2,
      ),
      dx.TableBorder(
        borderSide: dx.TableBorderSide.insideV,
        color: '3434DF',
        pbrStyle: dx.ParagraphBorderStyle.single,
        width: 2,
      ),
    ],
  ),
);

docx.attachTable(table);
```

Add comments to the document with attachComment, for example:

```dart
docx.attachComment(
  message: 'Chapter 5 is still under construction',
  author: 'The creator',
  initials: 'TC');
```

When the document is finished, create the file. This file is stored in the temporary directory of DocxBuilder. Process, send or save this file.

```dart
final File f = await docx.createDocXFile(documentTitle: 'My new document', documentCreator: 'Me');
```

If the file has been processed or if DocxBuilder is no longer needed, do not forget to delete the cache and free resources with clear:

```dart
docx.clear();
```

### Extra

DocXBuilder has a lot of options available for OOXML. It is impossible to explain them all on one page. Read the documentation and feel free to experiment with different settings and functions!
