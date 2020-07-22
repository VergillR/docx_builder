import 'dart:io';

import 'package:docx_builder/src/utils/utils.dart';
import 'package:docx_builder/src/styles/style_classes/index.dart';

import 'builders/index.dart' as _b;
import 'package/packager.dart' as _p;
import 'utils/constants/constants.dart' as _c;

export './styles/style_classes/index.dart';
export './styles/styles.dart';

/// DocXBuilder is used to construct and create .docx files.
///
/// Add global styling with setGlobalDocxTextStyle, setGlobalDocxPageStyle and setDocumentBackgroundColor.
///
/// Add text to the document with addText or addMixedText.
/// AddText uses the global styling.
/// AddMixedText overrides global styles with its own custom styling rules (if provided).
///
/// CreateDocXFile() is used to create the .docx file.
/// When done, clear() is used to destroy the cache.
class DocXBuilder {
  _p.Packager _packager;
  final StringBuffer _docxstring = StringBuffer();
  bool _bufferClosed = false;
  int _charCount = 0;
  int _charCountWithSpaces = 0;
  int _parCount = 0;

  String _documentBackgroundColor;
  String get documentBackgroundColor => _documentBackgroundColor;

  DocxTextStyle _globalDocxTextStyle;
  DocxTextStyle get globalDocxTextStyle => _globalDocxTextStyle;
  DocxPageStyle _globalDocxPageStyle;
  DocxPageStyle get globalDocxPageStyle => _globalDocxPageStyle;

  static String mimetype =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

  /// Assign a [cacheDirectory] that DocXBuilder can use to store temporary files.
  DocXBuilder(Directory cacheDirectory) {
    _packager = _p.Packager(cacheDirectory);
    _initDocX();
  }

  void _initDocX() {
    _bufferClosed = false;
    _docxstring.write(_c.startDoc);
  }

  /// Add a counter for characters; These values are written to the app.xml file.
  void _addToCharCounters(String text) {
    _charCountWithSpaces += text.length;
    _charCount += text.replaceAll(' ', '').length;
  }

  /// This attribute should be set before adding anything to the document or else it will be ignored.
  /// The given [color] should be in "RRGGBB" format.
  /// This is a global setting for the entire document which is placed before the body tag.
  void setDocumentBackgroundColor(String color) {
    if (!_bufferClosed &&
        _documentBackgroundColor == null &&
        isValidColor(color)) {
      _documentBackgroundColor = '<w:background w:color="$color" /><w:body>';
      if (_docxstring.toString().endsWith('<w:body>')) {
        _docxstring.clear();
        final String docWithBgColor =
            _c.startDoc.replaceAll('<w:body>', _documentBackgroundColor);
        _docxstring.write(docWithBgColor);
      }
    }
  }

  // Blank page can also be achieved with sectPr, sectType.nextPage and then create a new section which prompts a new page
  // void addBlankPageA4() => _docxstring.write(_c.blankPageA4);

  // ignore: use_setters_to_change_properties
  /// Sets the global page style, such as orientation, size, and page borders.
  /// This page style can only be set once as the final SectPr directly inside the body tag of the document.
  /// When adding custom section properties to a paragraph, make sure that paragraph is NOT the last paragraph of the document or else the document will be malformed.
  void setGlobalDocxPageStyle(DocxPageStyle pageStyle) =>
      _globalDocxPageStyle = pageStyle;

  /// Obtain the XML string of the page style.
  /// If no style is given, then the globalDocxPageStyle is used.
  String _getDocxPageStyleAsString({DocxPageStyle style}) {
    final DocxPageStyle pageStyle = style ?? _globalDocxPageStyle;
    return _b.SectPr.getSpr(
      cols: pageStyle.cols,
      colsHaveSeparator: pageStyle.colsHaveSeparator,
      colsHaveEqualWidth: pageStyle.colsHaveEqualWidth,
      sectType: pageStyle.sectType,
      pageSzHeight: pageStyle.pageSzHeight,
      pageSzWidth: pageStyle.pageSzWidth,
      pageOrientation: pageStyle.pageOrientation,
      sectVerticalAlign: pageStyle.sectVerticalAlign,
      pageBorders: pageStyle.pageBorders,
      pageBorderDisplay: pageStyle.pageBorderDisplay,
      pageBorderOffsetBasedOnText: pageStyle.pageBorderOffsetBasedOnText,
      pageBorderIsRenderedAboveText: pageStyle.pageBorderIsRenderedAboveText,
      pageNumberingFormat: pageStyle.pageNumberingFormat,
      pageNumberingStart: pageStyle.pageNumberingStart,
      lineNumbering: pageStyle.lineNumbering,
      pageMargin: pageStyle.pageMargin,
    );
  }

  // ignore: use_setters_to_change_properties
  /// Sets the global style for text. Except for HighlightColor, colors are always 'RRGGBB' format.
  /// The global text style can be overridden by addMixedText with its own styling rules.
  void setGlobalDocxTextStyle(DocxTextStyle textStyle) =>
      _globalDocxTextStyle = textStyle;

  /// Obtain the XML string of the text style.
  /// If no style is given, then the globalDocxTextStyle is used.
  String _getTextStyleAsString({DocxTextStyle style}) {
    final DocxTextStyle textStyle = style ?? _globalDocxTextStyle;
    return _b.Ppr.getPpr(
      rprString: _b.Rpr.getRpr(
        bold: textStyle.bold,
        caps: textStyle.caps,
        doubleStrike: textStyle.doubleStrike,
        fontColor: textStyle.fontColor,
        fontName: textStyle.fontName,
        fontNameComplexScript: textStyle.fontNameComplexScript,
        fontSize: textStyle.fontSize,
        highlightColor: textStyle.highlightColor,
        italic: textStyle.italic,
        rtlText: textStyle.rtlText,
        shading: textStyle.shading,
        smallCaps: textStyle.smallCaps,
        strike: textStyle.strike,
        textArt: textStyle.textArt,
        underline: textStyle.underline,
        underlineColor: textStyle.underlineColor,
        vanish: textStyle.vanish,
        vertAlign: textStyle.vertAlign,
        spacing: textStyle.textSpacing,
        stretchOrCompressInPercentage: textStyle.stretchOrCompressInPercentage,
        kern: textStyle.kern,
        fitText: textStyle.fitText,
        effect: textStyle.textEffect,
      ),
      textAlignment: textStyle.textAlignment,
      paragraphBorderOnAllSides: textStyle.paragraphBorderOnAllSides,
      paragraphBorders: textStyle.paragraphBorders,
      paragraphIndent: textStyle.paragraphIndent,
      paragraphShading: textStyle.paragraphShading,
      spacing: textStyle.paragraphSpacing,
      tabs: textStyle.tabs,
      vAlign: textStyle.vAlign,
      keepLines: textStyle.keepLines,
      keepNext: textStyle.keepNext,
    );
  }

  /// AddText adds lines of text to the document.
  /// It uses the global text styling defined by setGlobalDocxTextStyle.
  /// This function always adds a new paragraph to the document.
  void addText(String text) {
    if (!_bufferClosed) {
      _docxstring.writeAll(<String>[
        '<w:p>',
        '<w:r>${_getTextStyleAsString()}${text.startsWith(' ') || text.endsWith(' ') ? '<w:t xml:space="preserve">' : '<w:t>'}$text</w:t></w:r></w:p>'
      ]);
      _addToCharCounters(text);
      _parCount++;
    }
  }

  /// AddMixedText adds lines of text that do NOT have the same styling as each other or with the global text style.
  /// Given lists should have equal lengths. Page style is optional and is used for custom section styling.
  /// Note that the last paragraph of the document should NOT have any custom section styling!
  /// Except for HighlightColor, colors are always 'RRGGBB' format.
  ///
  /// Lists can hold null values. For text: null implies no text; for textStyles: null implies use of globalDocxTextStyle.
  ///
  /// This function always adds a new paragraph to the document.
  void addMixedText(List<String> text, List<DocxTextStyle> textStyles,
      {DocxPageStyle pageStyle}) {
    if (!_bufferClosed && text.isNotEmpty && text.length == textStyles.length) {
      _docxstring.write('<w:p>');
      if (pageStyle != null) {
        _docxstring.write(_getDocxPageStyleAsString(style: pageStyle));
      }
      for (int i = 0; i < text.length; i++) {
        final String t = text[i] ?? '';
        _docxstring.writeAll(<String>[
          '<w:r>',
          _getTextStyleAsString(style: textStyles[i]),
          '<w:t xml:space="preserve">$t</w:t></w:r>',
        ]);
        _addToCharCounters(t);
      }
      _docxstring.write('</w:p>');
      _parCount++;
    }
  }

  /// InsertImage inserts an image at the current position in the buffer.
  void insertImage() {
    if (!_bufferClosed) {}
  }

  /// Defines a table that can be inserted later.
  Table defineTable() {
    if (!_bufferClosed) {}
    return Table();
  }

  // InsertTable inserts a table that was defined earlier with defineTable.
  void insertTable(Table table) {
    if (!_bufferClosed) {}
  }

  /// Create the .docx file with the content stored in the buffer of DocXBuilder.
  /// This also closes the buffer.
  /// After processing the .docx file, do not forget to call clear to free resources and reopen the buffer.
  Future<File> createDocXFile() async {
    try {
      if (!_bufferClosed) {
        final String lastSectPr = _getDocxPageStyleAsString(
            style:
                _globalDocxPageStyle ?? DocxPageStyle().getDefaultPageStyle());
        _docxstring.write(lastSectPr);
        _docxstring.write('</w:body></w:document>');
        _bufferClosed = true;
      }

      final File f = await _packager.bundle(
        _docxstring.toString(),
        chars: _charCount,
        charsWithSpaces: _charCountWithSpaces,
        paragraphs: _parCount,
      );
      return f;
    } catch (e) {
      throw StateError('DOCX BUILDER ERROR: $e');
    }
  }

  /// Deletes all data and cache, so a new document can be created (or when DocXBuilder is no longer needed).
  /// This function also destroys the source file produced by createDocxFile so be sure to save or process the source file before calling clear.
  void clear({bool resetDocX = true}) {
    _docxstring.clear();
    _documentBackgroundColor = null;
    _globalDocxPageStyle = null;
    _globalDocxTextStyle = null;
    _packager.destroyCache();
    if (resetDocX) {
      _initDocX();
    }
  }
}
