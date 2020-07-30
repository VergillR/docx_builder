import 'dart:io';

import 'package:meta/meta.dart';
import 'package:docx_builder/src/utils/constants/mimetypes.dart';
import 'package:docx_builder/src/utils/utils.dart';
import 'package:docx_builder/src/styles/style_classes/index.dart';
import 'package:docx_builder/src/styles/style_containers/index.dart';
import 'package:docx_builder/src/styles/style_enums.dart';

import 'builders/index.dart' as _b;
import 'package/packager.dart' as _p;
import 'utils/constants/constants.dart' as _c;

export './styles/style_classes/index.dart';
export './styles/style_containers/index.dart';
export 'styles/style_enums.dart';

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
  final StringBuffer _docx = StringBuffer();
  bool _bufferClosed = false;
  // int _charCount = 0;
  // int _charCountWithSpaces = 0;
  // int _parCount = 0;
  int _headerCounter = 1;
  int _footerCounter = 1;

  String _documentBackgroundColor;
  String get documentBackgroundColor => _documentBackgroundColor;

  TextStyle _globalTextStyle = TextStyle();
  TextStyle get globalTextStyle => _globalTextStyle;
  PageStyle _globalPageStyle = PageStyle().getDefaultPageStyle();
  PageStyle get globalPageStyle => _globalPageStyle;

  Header _firstPageHeader;
  Header _oddPageHeader;
  Header _evenPageHeader;
  Footer _firstPageFooter;
  Footer _oddPageFooter;
  Footer _evenPageFooter;
  bool _insertHeadersAndFootersInThisSection = false;

  final String mimetype =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

  final int emuWidthA4 = 7560000;
  final int emuHeightA4 = 10692000;
  final int twipsWidthA4 = 11906;
  final int twipsHeightA4 = 16838;

  final String bullet = '•';
  final String arrowBullet = '‣';
  final String hyphenBullet = '⁃';
  final String inverseBullet = '◘';
  final String openBullet = '◦';

  /// Convert inches to Twips (twentieth of point).
  int convertInchesToTwips(int inches) => inches * 1440;

  /// Convert centimeters to Twips (twentieth of point).
  int convertCmToTwips(int cm) => (cm * 566.92913386).round();

  /// Convert millimeters to Twips (twentieth of point).
  int convertMmToTwips(int mm) => (mm * 56.692913386).round();

  /// Convert inches to EMU; A4 format is 8.3 x 11.7 inches.
  int convertInchesToEMU(int inches) => inches * 914400;

  /// Convert centimeters to EMU; A4 format is 21x29.7cm.
  int convertCentimetersToEMU(int cm) => cm * 360000;

  /// Convert millimeters to EMU; A4 format is 210x297mm.
  int convertMillimetersToEMU(int mm) => mm * 36000;

  /// Assign a [cacheDirectory] that DocXBuilder can use to store temporary files.
  DocXBuilder(Directory cacheDirectory) {
    _packager = _p.Packager(cacheDirectory);
    _initDocX();
  }

  void _initDocX() {
    _bufferClosed = false;
    _docx.write(_c.startDoc);
  }

  /// Add a counter for characters; These values are written to the app.xml file.
  // void _addToCharCounters(String text) {
  //   _charCountWithSpaces += text.length;
  //   _charCount += text.replaceAll(' ', '').length;
  // }

  /// This attribute should be set before adding anything to the document or else it will be ignored.
  /// The given [color] should be in "RRGGBB" format.
  /// This is a global setting for the entire document which is placed before the body tag.
  void setDocumentBackgroundColor(String color) {
    if (!_bufferClosed &&
        _documentBackgroundColor == null &&
        isValidColor(color)) {
      _documentBackgroundColor = '<w:background w:color="$color" /><w:body>';
      if (_docx.toString().endsWith('<w:body>')) {
        _docx.clear();
        final String docWithBgColor =
            _c.startDoc.replaceAll('<w:body>', _documentBackgroundColor);
        _docx.write(docWithBgColor);
      }
    }
  }

  // Blank page can also be achieved with sectPr, sectType.nextPage and then create a new section which prompts a new page
  // void addBlankPageA4() => _docx.write(_c.blankPageA4);

  // ignore: use_setters_to_change_properties
  /// Sets the global page style, such as orientation, size, and page borders.
  /// This page style can only be set once as the final SectPr directly inside the body tag of the document.
  /// When adding custom section properties to a paragraph, make sure that paragraph is NOT the last paragraph of the document or else the document will be malformed.
  void setGlobalPageStyle(PageStyle pageStyle) => _globalPageStyle = pageStyle;

  /// The [type] determines where the header will appear: for odd pages (also called default header), first page (also called title page header) or even pages. When a header is set, it cannot be changed or removed afterwards as additional files and references are immediately written to the cache directory upon creation.
  ///
  /// Headers are not visible until they have been attached to the document by calling attachHeadersAndFooters.
  ///
  /// Make sure the lists [text] and [textStyles] have the same length or else the textStyles will be treated as null.
  ///
  /// If given, [complexFields] should have the same length as [text] and includes complex fields, such as page and date. For example, ComplexField(instructions: 'PAGE') instructs the word processor to insert the current page number. A complexField cannot be added if hyperlinkTo is not null.
  void setHeader(
    HeaderType headerType,
    List<String> text,
    List<TextStyle> textStyles, {
    bool doNotUseGlobalTextStyle = true,
    List<ComplexField> complexFields,
  }) {
    if (!_bufferClosed) {
      final List<TextStyle> styles = text.length == textStyles.length
          ? textStyles
          : List<TextStyle>.generate(text.length, (index) => null);
      _initHeaderOrFooter(
        headerType: headerType,
        text: text,
        textStyles: styles,
        doNotUseGlobalTextStyle: doNotUseGlobalTextStyle,
        complexFields: complexFields,
      );
    }
  }

  /// The [type] determines where the footer will appear: for odd pages (also called default footer), first page (also called title page footer) or even pages. When a footer is set, it cannot be changed or removed afterwards as additional files and references are immediately written to the cache directory upon creation.
  ///
  /// Footers are not visible until they have been attached to the document by calling attachHeadersAndFooters.
  ///
  /// Make sure the lists [text] and [textStyles] have the same length or else the textStyles will be treated as null.
  ///
  /// If given, [complexFields] should have the same length as [text] and includes complex fields, such as page and date. For example, ComplexField(instructions: 'PAGE') instructs the word processor to insert the current page number. A complexField cannot be added if hyperlinkTo is not null.
  void setFooter(
    FooterType footerType,
    List<String> text,
    List<TextStyle> textStyles, {
    bool doNotUseGlobalTextStyle = true,
    List<ComplexField> complexFields,
  }) {
    if (!_bufferClosed) {
      final List<TextStyle> styles = text.length == textStyles.length
          ? textStyles
          : List<TextStyle>.generate(text.length, (index) => null);
      _initHeaderOrFooter(
        footerType: footerType,
        text: text,
        textStyles: styles,
        doNotUseGlobalTextStyle: doNotUseGlobalTextStyle,
        complexFields: complexFields,
      );
    }
  }

  /// Attaches all given headers and footers to the document by marking them to be included in the next SectionProperties.
  /// Only when attached, are headers and footers actually visible in the document.
  /// Preferably, call this function directly after all necessary headers and footers have been set.
  ///
  /// Normally, this function should only be called once.
  ///
  /// Calling it multiple times with different sets of headers/footers is possible but is considered to be experimental.
  void attachHeadersAndFooters() {
    if (!_insertHeadersAndFootersInThisSection) {
      _insertHeadersAndFootersInThisSection = true;
    }
  }

  void _initHeaderOrFooter(
      {HeaderType headerType,
      FooterType footerType,
      List<String> text,
      List<TextStyle> textStyles,
      List<ComplexField> complexFields,
      bool doNotUseGlobalTextStyle = false}) {
    final bool isHeader = headerType != null;

    final StringBuffer b = StringBuffer();
    if (text.isNotEmpty && text.length == textStyles.length) {
      b.write('<w:p><w:pPr>');
      b.write(_getParagraphStyleAsString(doNotUseGlobalStyle: true)
          .replaceFirst('<w:pPr>', ''));

      final List<ComplexField> cf =
          complexFields != null && complexFields.length == text.length
              ? complexFields
              : List<ComplexField>.generate(text.length, (index) => null);

      for (int i = 0; i < text.length; i++) {
        final String t = text[i] ?? '';
        final ComplexField f = cf[i];
        if (f == null || f.instructions == null || f.includeSeparate == null) {
          String closeHyperlink = '';
          String openHyperlink = '';
          if (textStyles[i] != null &&
              textStyles[i].hyperlinkTo != null &&
              textStyles[i].hyperlinkTo.isNotEmpty) {
            _packager.addHyperlink(textStyles[i].hyperlinkTo);
            closeHyperlink = '</w:hyperlink>';
            openHyperlink = '<w:hyperlink r:id="rId${_packager.rIdCount - 1}">';
          }
          b.writeAll(<String>[
            openHyperlink,
            '<w:r>',
            _getTextStyleAsString(
                styleAsHyperlink: textStyles[i] != null &&
                    textStyles[i].hyperlinkTo != null &&
                    textStyles[i].hyperlinkTo.isNotEmpty,
                style: textStyles[i],
                doNotUseGlobalStyle: doNotUseGlobalTextStyle),
            '<w:t xml:space="preserve">$t</w:t></w:r>$closeHyperlink',
          ]);
        } else {
          b.writeAll(<String>[
            '<w:r>',
            _getTextStyleAsString(
                style: textStyles[i],
                doNotUseGlobalStyle: doNotUseGlobalTextStyle),
            '<w:t xml:space="preserve">$t</w:t></w:r><w:r>${_getTextStyleAsString(style: textStyles[i], doNotUseGlobalStyle: doNotUseGlobalTextStyle)}${f.getXml()}</w:r>',
          ]);
        }
        // _addToCharCounters(t);
      }

      b.write('</w:p>');
      // _parCount++;
    }

    final int counter = isHeader ? _headerCounter++ : _footerCounter++;
    bool allowEvenPages = false;
    final String contents = b.toString();

    if (isHeader) {
      if (headerType == HeaderType.firstPage) {
        _firstPageHeader = FirstPageHeader(
            rId: 'rId${_packager.rIdCount}', taggedText: contents);
      } else if (headerType == HeaderType.oddPage) {
        _oddPageHeader = OddPageHeader(
            rId: 'rId${_packager.rIdCount}', taggedText: contents);
      } else {
        allowEvenPages = true;
        _evenPageHeader = EvenPageHeader(
            rId: 'rId${_packager.rIdCount}', taggedText: contents);
      }
    } else {
      if (footerType == FooterType.firstPage) {
        _firstPageFooter = FirstPageFooter(
            rId: 'rId${_packager.rIdCount}', taggedText: contents);
      } else if (footerType == FooterType.oddPage) {
        _oddPageFooter = OddPageFooter(
            rId: 'rId${_packager.rIdCount}', taggedText: contents);
      } else {
        allowEvenPages = true;
        _evenPageFooter = EvenPageFooter(
            rId: 'rId${_packager.rIdCount}', taggedText: contents);
      }
    }
    _packager.addHeaderOrFooter(counter, contents,
        isHeader: isHeader, evenPage: allowEvenPages);
  }

  /// Should only be called once after all headers and footers are set, to insert the references of headers and footers to a SectPr segment in the document.
  String _addHeadersAndFootersToSectPr(String sectPr) {
    if (_firstPageHeader != null ||
        _firstPageFooter != null ||
        _oddPageHeader != null ||
        _oddPageFooter != null ||
        _evenPageHeader != null ||
        _evenPageFooter != null) {
      final List<Header> h = <Header>[
        _firstPageHeader,
        _oddPageHeader,
        _evenPageHeader
      ];
      final List<Footer> f = <Footer>[
        _firstPageFooter,
        _oddPageFooter,
        _evenPageFooter
      ];

      bool titlePageHasOwnHeaderFooter = false;
      final StringBuffer hf = StringBuffer();

      for (int i = 0; i < h.length; i++) {
        if (h[i] != null) {
          if (i == 0) {
            titlePageHasOwnHeaderFooter = true;
          }
          hf.write(
              '<w:headerReference r:id="${h[i].rId}" w:type="${i == 0 ? "first" : i == 1 ? "default" : "even"}"/>');
        }
      }
      for (int i = 0; i < f.length; i++) {
        if (f[i] != null) {
          if (i == 0) {
            titlePageHasOwnHeaderFooter = true;
          }
          hf.write(
              '<w:footerReference r:id="${f[i].rId}" w:type="${i == 0 ? "first" : i == 1 ? "default" : "even"}"/>');
        }
      }
      if (titlePageHasOwnHeaderFooter) {
        hf.write('<w:titlePg/>');
      }
      return sectPr.replaceFirst('<w:sectPr>', '<w:sectPr>${hf.toString()}');
    } else {
      return '';
    }
  }

  /// Obtain the XML string of the page style.
  /// If no style is given, then the globalDocxPageStyle is used.
  String _getPageStyleAsString(
      {PageStyle style, bool doNotUseGlobalStyle = false}) {
    final PageStyle pageStyle = style ?? PageStyle();
    String spr = _b.SectPr.getSpr(
      cols: doNotUseGlobalStyle
          ? pageStyle.cols
          : pageStyle.cols ?? _globalPageStyle.cols,
      colsHaveSeparator: doNotUseGlobalStyle
          ? pageStyle.colsHaveSeparator
          : pageStyle.colsHaveSeparator ?? _globalPageStyle.colsHaveSeparator,
      colsHaveEqualWidth: doNotUseGlobalStyle
          ? pageStyle.colsHaveEqualWidth
          : pageStyle.colsHaveEqualWidth ?? _globalPageStyle.colsHaveEqualWidth,
      sectType: doNotUseGlobalStyle
          ? pageStyle.sectType
          : pageStyle.sectType ?? _globalPageStyle.sectType,
      pageSzHeight: doNotUseGlobalStyle
          ? pageStyle.pageSzHeight
          : pageStyle.pageSzHeight ?? _globalPageStyle.pageSzHeight,
      pageSzWidth: doNotUseGlobalStyle
          ? pageStyle.pageSzWidth
          : pageStyle.pageSzWidth ?? _globalPageStyle.pageSzWidth,
      pageOrientation: doNotUseGlobalStyle
          ? pageStyle.pageOrientation
          : pageStyle.pageOrientation ?? _globalPageStyle.pageOrientation,
      sectVerticalAlign: doNotUseGlobalStyle
          ? pageStyle.sectVerticalAlign
          : pageStyle.sectVerticalAlign ?? _globalPageStyle.sectVerticalAlign,
      pageBorders: doNotUseGlobalStyle
          ? pageStyle.pageBorders
          : pageStyle.pageBorders ?? _globalPageStyle.pageBorders,
      pageBorderDisplay: doNotUseGlobalStyle
          ? pageStyle.pageBorderDisplay
          : pageStyle.pageBorderDisplay ?? _globalPageStyle.pageBorderDisplay,
      pageBorderOffsetBasedOnText: doNotUseGlobalStyle
          ? pageStyle.pageBorderOffsetBasedOnText
          : pageStyle.pageBorderOffsetBasedOnText ??
              _globalPageStyle.pageBorderOffsetBasedOnText,
      pageBorderIsRenderedAboveText: doNotUseGlobalStyle
          ? pageStyle.pageBorderIsRenderedAboveText
          : pageStyle.pageBorderIsRenderedAboveText ??
              _globalPageStyle.pageBorderIsRenderedAboveText,
      pageNumberingFormat: doNotUseGlobalStyle
          ? pageStyle.pageNumberingFormat
          : pageStyle.pageNumberingFormat ??
              _globalPageStyle.pageNumberingFormat,
      pageNumberingStart: doNotUseGlobalStyle
          ? pageStyle.pageNumberingStart
          : pageStyle.pageNumberingStart ?? _globalPageStyle.pageNumberingStart,
      lineNumbering: doNotUseGlobalStyle
          ? pageStyle.lineNumbering
          : pageStyle.lineNumbering ?? _globalPageStyle.lineNumbering,
      pageMargin: doNotUseGlobalStyle
          ? pageStyle.pageMargin
          : pageStyle.pageMargin ?? _globalPageStyle.pageMargin,
    );
    if (_insertHeadersAndFootersInThisSection) {
      spr = _addHeadersAndFootersToSectPr(spr);
      _insertHeadersAndFootersInThisSection = false;
    }
    return spr;
  }

  // ignore: use_setters_to_change_properties
  /// Sets the global style for text. Except for HighlightColor, colors are in 'RRGGBB' format.
  /// The global text style can be overridden by addMixedText with its own styling rules.
  /// The hyperlinkTo and textFrame properties of the global text style are never used as this makes no sense. Hyperlinks and textframes can be added manually with addText or addMixedText.
  void setGlobalTextStyle(TextStyle textStyle) => _globalTextStyle = textStyle;

  /// Obtain the XML string of the paragraph style, such as text alignment.
  /// If no style is given, then the globalDocxTextStyle is used (unless [doNotUseGlobalStyle] is set to true).
  /// Textframes cannot be set global.
  String _getParagraphStyleAsString(
      {TextStyle textStyle, bool doNotUseGlobalStyle = false}) {
    final TextStyle style = textStyle ?? TextStyle();
    return _b.Ppr.getPpr(
      textFrame: textStyle?.textFrame != null ? textStyle.textFrame : null,
      keepLines: doNotUseGlobalStyle
          ? style.keepLines
          : style.keepLines ?? _globalTextStyle.keepLines,
      keepNext: doNotUseGlobalStyle
          ? style.keepNext
          : style.keepNext ?? _globalTextStyle.keepNext,
      paragraphBorderOnAllSides: doNotUseGlobalStyle
          ? style.paragraphBorderOnAllSides
          : style.paragraphBorderOnAllSides ??
              _globalTextStyle.paragraphBorderOnAllSides,
      paragraphBorders: doNotUseGlobalStyle
          ? style.paragraphBorders
          : style.paragraphBorders ?? _globalTextStyle.paragraphBorders,
      paragraphIndent: doNotUseGlobalStyle
          ? style.paragraphIndent
          : style.paragraphIndent ?? _globalTextStyle.paragraphIndent,
      paragraphShading: doNotUseGlobalStyle
          ? style.paragraphShading
          : style.paragraphShading ?? _globalTextStyle.paragraphShading,
      spacing: doNotUseGlobalStyle
          ? style.paragraphSpacing
          : style.paragraphSpacing ?? _globalTextStyle.paragraphSpacing,
      tabs: doNotUseGlobalStyle
          ? style.tabs
          : style.tabs ?? _globalTextStyle.tabs,
      textAlignment: doNotUseGlobalStyle
          ? style.textAlignment
          : style.textAlignment ?? _globalTextStyle.textAlignment,
      verticalTextAlignment: doNotUseGlobalStyle
          ? style.verticalTextAlignment
          : style.verticalTextAlignment ??
              _globalTextStyle.verticalTextAlignment,
    );
  }

  /// Obtain the XML string of the text style.
  /// If no style is given, then the globalDocxTextStyle is used (unless [doNotUseGlobalStyle] is set to true).
  String _getTextStyleAsString(
      {TextStyle style,
      bool doNotUseGlobalStyle = false,
      bool styleAsHyperlink = false}) {
    final TextStyle textStyle = style ?? TextStyle();
    return _b.Rpr.getRpr(
      styleAsHyperlink: styleAsHyperlink,
      bold: doNotUseGlobalStyle
          ? textStyle.bold
          : textStyle.bold ?? _globalTextStyle.bold,
      caps: doNotUseGlobalStyle
          ? textStyle.caps
          : textStyle.caps ?? _globalTextStyle.caps,
      doubleStrike: doNotUseGlobalStyle
          ? textStyle.doubleStrike
          : textStyle.doubleStrike ?? _globalTextStyle.doubleStrike,
      fontColor: doNotUseGlobalStyle
          ? textStyle.fontColor
          : textStyle.fontColor ?? _globalTextStyle.fontColor,
      fontName: doNotUseGlobalStyle
          ? textStyle.fontName
          : textStyle.fontName ?? _globalTextStyle.fontName,
      fontNameComplexScript: doNotUseGlobalStyle
          ? textStyle.fontNameComplexScript
          : textStyle.fontNameComplexScript ??
              _globalTextStyle.fontNameComplexScript,
      fontSize: doNotUseGlobalStyle
          ? textStyle.fontSize
          : textStyle.fontSize ?? _globalTextStyle.fontSize,
      highlightColor: doNotUseGlobalStyle
          ? textStyle.highlightColor
          : textStyle.highlightColor ?? _globalTextStyle.highlightColor,
      italic: doNotUseGlobalStyle
          ? textStyle.italic
          : textStyle.italic ?? _globalTextStyle.italic,
      rtlText: doNotUseGlobalStyle
          ? textStyle.rtlText
          : textStyle.rtlText ?? _globalTextStyle.rtlText,
      shading: doNotUseGlobalStyle
          ? textStyle.shading
          : textStyle.shading ?? _globalTextStyle.shading,
      smallCaps: doNotUseGlobalStyle
          ? textStyle.smallCaps
          : textStyle.smallCaps ?? _globalTextStyle.smallCaps,
      strike: doNotUseGlobalStyle
          ? textStyle.strike
          : textStyle.strike ?? _globalTextStyle.strike,
      textArt: doNotUseGlobalStyle
          ? textStyle.textArt
          : textStyle.textArt ?? _globalTextStyle.textArt,
      underline: doNotUseGlobalStyle
          ? textStyle.underline
          : textStyle.underline ?? _globalTextStyle.underline,
      underlineColor: doNotUseGlobalStyle
          ? textStyle.underlineColor
          : textStyle.underlineColor ?? _globalTextStyle.underlineColor,
      vanish: doNotUseGlobalStyle
          ? textStyle.vanish
          : textStyle.vanish ?? _globalTextStyle.vanish,
      vertAlign: doNotUseGlobalStyle
          ? textStyle.vertAlign
          : textStyle.vertAlign ?? _globalTextStyle.vertAlign,
      spacing: doNotUseGlobalStyle
          ? textStyle.textSpacing
          : textStyle.textSpacing ?? _globalTextStyle.textSpacing,
      stretchOrCompressInPercentage: doNotUseGlobalStyle
          ? textStyle.stretchOrCompressInPercentage
          : textStyle.stretchOrCompressInPercentage ??
              _globalTextStyle.stretchOrCompressInPercentage,
      kern: doNotUseGlobalStyle
          ? textStyle.kern
          : textStyle.kern ?? _globalTextStyle.kern,
      fitText: doNotUseGlobalStyle
          ? textStyle.fitText
          : textStyle.fitText ?? _globalTextStyle.fitText,
      effect: doNotUseGlobalStyle
          ? textStyle.textEffect
          : textStyle.textEffect ?? _globalTextStyle.textEffect,
    );
  }

  /// AddText adds lines of text to the document.
  /// It uses the global text styling defined by setGlobalDocxTextStyle.
  /// This function always adds a new paragraph to the document.
  ///
  /// If [lineOrPageBreak] is given, then a LineBreak will be added at the end of the text.
  /// If globalDocxTextStyle has a non-empty Tabs list, then a tab can be added in front of the text by setting [addTab] to true.
  ///
  /// If [hyperlinkTo] is not null or empty, then the text will be a hyperlink and direct to [hyperlinkTo]. The global textstyle's hyperlinkTo is always ignored.
  /// If given, [complexField] adds a complex field, such as page and date, e.g. ComplexField(instructions: 'PAGE') instructs the word processor to insert the current page number. A complexField cannot be added if hyperlinkTo is not null.
  void addText(
    String text, {
    LineBreak lineOrPageBreak,
    bool addTab = false,
    String hyperlinkTo,
    ComplexField complexField,
    TextFrame textFrame,
  }) {
    if (!_bufferClosed) {
      _docx.write(_getCachedAddText(
        text,
        lineOrPageBreak: lineOrPageBreak,
        addTab: addTab,
        hyperlinkTo: hyperlinkTo,
        complexField: complexField,
        textFrame: textFrame,
      ));
      // _addToCharCounters(text);
      // _parCount++;
    }
  }

  /// When the table has been finalized, attach it to the document at the cursor's current position. A table should contain a list of column widths and have at least 1 table row. Rows can contain zero or more table cells.
  ///
  /// Similar to headers and footers, tables will only become visible after being attached to the document.
  void attachTable(Table table) {
    if (!_bufferClosed) {
      _docx.write(table.getXml());
    }
  }

  /// Assigning cells to tableRow.tableCells directly is also possible.
  void insertTableCellsInRow(
          {@required List<TableCell> tableCells,
          @required TableRow tableRow}) =>
      tableRow.tableCells = tableCells;

  /// Assigning rows to table.tableRows directly is also possible.
  void insertRowsInTable(
          {@required Table table, @required List<TableRow> tableRows}) =>
      table.tableRows = tableRows;

  void insertTableInTableCell(
          {@required Table table, @required TableCell tableCell}) =>
      tableCell.xmlContent = table.getXml();

  void insertTextInTableCell({
    @required TableCell tableCell,
    @required String text,
    LineBreak lineOrPageBreak,
    bool addTab = false,
    String hyperlinkTo,
    ComplexField complexField,
    TextFrame textFrame,
  }) {
    if (!_bufferClosed) {
      tableCell.xmlContent = _getCachedAddText(
        text,
        lineOrPageBreak: lineOrPageBreak,
        addTab: addTab,
        hyperlinkTo: hyperlinkTo,
        complexField: complexField,
        textFrame: textFrame,
      );
      // _addToCharCounters(text);
      // _parCount++;
    }
  }

  /// Writes XML text to cache; This data can then either be written to a stringbuffer or used as content for table cells.
  String _getCachedAddText(
    String text, {
    LineBreak lineOrPageBreak,
    bool addTab = false,
    String hyperlinkTo,
    ComplexField complexField,
    TextFrame textFrame,
  }) {
    final StringBuffer cached = StringBuffer();
    final String tab =
        globalTextStyle.tabs != null && addTab ? '<w:r><w:tab/></w:r>' : '';
    final String lineBreak =
        lineOrPageBreak != null ? lineOrPageBreak.getXml() : '';

    final String ppr = textFrame == null
        ? _getParagraphStyleAsString()
        : _getParagraphStyleAsString()
            .replaceFirst('</w:pPr>', '${textFrame.getXml()}</w:pPr>');

    if (complexField == null ||
        complexField.instructions == null ||
        complexField.includeSeparate == null) {
      String closeHyperlink = '';
      String openHyperlink = '';
      if (hyperlinkTo != null && hyperlinkTo.isNotEmpty) {
        _packager.addHyperlink(hyperlinkTo);
        closeHyperlink = '</w:hyperlink>';
        openHyperlink = '<w:hyperlink r:id="rId${_packager.rIdCount - 1}">';
      }

      cached.writeAll(<String>[
        '<w:p>$ppr',
        tab,
        openHyperlink,
        '<w:r>${_getTextStyleAsString(styleAsHyperlink: openHyperlink.isNotEmpty)}${text.startsWith(' ') || text.endsWith(' ') ? '<w:t xml:space="preserve">' : '<w:t>'}$text</w:t></w:r>$lineBreak$closeHyperlink</w:p>'
      ]);
    } else {
      cached.writeAll(<String>[
        '<w:p>$ppr',
        tab,
        '<w:r>${_getTextStyleAsString()}<w:t xml:space="preserve">$text</w:t></w:r><w:r>${_getTextStyleAsString()}${complexField.getXml()}</w:r>$lineBreak</w:p>'
      ]);
    }
    return cached.toString();
  }

  /// AddMixedText adds lines of text that do NOT have the same styling as each other or with the global text style.
  /// Unless [doNotUseTextGlobalStyle] is set to true, global text styling will be used for any values not provided (i.e. null values) by custom styling rules.
  /// Given lists should have equal lengths. Page style is optional and is used for custom section styling. If [doNotUsePageGlobalStyle] is set to false, global page styling will be used for any values not provided (i.e. null values) by the custom page styling.
  /// Note that the last paragraph of the document should NOT have any custom section styling!
  /// Except for HighlightColor, colors are in 'RRGGBB' format.
  ///
  /// Lists can hold null values. For text: null implies no text; for textStyles: null implies use of globalDocxTextStyle.
  ///
  /// This function always adds a new paragraph to the document.
  ///
  /// If [lineOrPageBreak] is given, then a LineBreak will be added after every item (if [addBreakAfterEveryItem] is true) or only after the last item on the [text] list (if [addBreakAfterEveryItem] is false, which is default).
  /// If globalDocxTextStyle has a non-empty Tabs list, then a tab can be added in front of the first text item by setting [addTab] to true.
  ///
  /// If the custom textstyles contain a hyperlinkTo value that is not null or empty, then the text will be a hyperlink and direct to [hyperlinkTo]. The global textstyle's hyperlinkTo is always ignored.
  ///
  /// If given, [complexFields] should have the same length as [text] and includes complex fields, such as page and date. For example, ComplexField(instructions: 'PAGE') instructs the word processor to insert the current page number. A complexField cannot be added if hyperlinkTo is not null.
  void addMixedText(
    List<String> text,
    List<TextStyle> textStyles, {
    PageStyle pageStyle,
    bool doNotUseGlobalTextStyle = false,
    bool doNotUseGlobalPageStyle = true,
    LineBreak lineOrPageBreak,
    bool addBreakAfterEveryItem = false,
    bool addTab = false,
    List<ComplexField> complexFields,
  }) {
    if (!_bufferClosed && text.isNotEmpty && text.length == textStyles.length) {
      _docx.write(_getCachedAddMixedText(
        text,
        textStyles,
        pageStyle: pageStyle,
        doNotUseGlobalTextStyle: doNotUseGlobalTextStyle,
        doNotUseGlobalPageStyle: doNotUseGlobalPageStyle,
        lineOrPageBreak: lineOrPageBreak,
        addBreakAfterEveryItem: addBreakAfterEveryItem,
        addTab: addTab,
        complexFields: complexFields,
      ));
      // _addToCharCounters(t);
      // _parCount++;
    }
  }

  void insertMixedTextInTableCell({
    @required TableCell tableCell,
    @required List<String> text,
    List<TextStyle> textStyles,
    PageStyle pageStyle,
    bool doNotUseGlobalTextStyle = false,
    bool doNotUseGlobalPageStyle = true,
    LineBreak lineOrPageBreak,
    bool addBreakAfterEveryItem = false,
    bool addTab = false,
    List<ComplexField> complexFields,
  }) {
    if (!_bufferClosed && text.isNotEmpty && text.length == textStyles.length) {
      tableCell.xmlContent = _getCachedAddMixedText(
        text,
        textStyles,
        pageStyle: pageStyle,
        doNotUseGlobalTextStyle: doNotUseGlobalTextStyle,
        doNotUseGlobalPageStyle: doNotUseGlobalPageStyle,
        lineOrPageBreak: lineOrPageBreak,
        addBreakAfterEveryItem: addBreakAfterEveryItem,
        addTab: addTab,
        complexFields: complexFields,
      );
      // _addToCharCounters(t);
      // _parCount++;
    }
  }

  /// Writes XML mixed text to cache; This data can then either be written to a stringbuffer or used as content for table cells.
  String _getCachedAddMixedText(
    List<String> text,
    List<TextStyle> textStyles, {
    PageStyle pageStyle,
    bool doNotUseGlobalTextStyle = false,
    bool doNotUseGlobalPageStyle = true,
    LineBreak lineOrPageBreak,
    bool addBreakAfterEveryItem = false,
    bool addTab = false,
    List<ComplexField> complexFields,
  }) {
    final StringBuffer d = StringBuffer();
    d.write('<w:p><w:pPr>');
    if (pageStyle != null) {
      d.write(_getPageStyleAsString(
          style: pageStyle, doNotUseGlobalStyle: doNotUseGlobalPageStyle));
    }
    d.write(_getParagraphStyleAsString().replaceFirst('<w:pPr>', ''));

    final String multiBreak = lineOrPageBreak != null && addBreakAfterEveryItem
        ? lineOrPageBreak.getXml()
        : '';

    if (globalTextStyle.tabs != null && addTab) {
      d.write('<w:r><w:tab/></w:r>');
    }

    final List<ComplexField> cf =
        complexFields != null && complexFields.length == text.length
            ? complexFields
            : List<ComplexField>.generate(text.length, (index) => null);

    for (int i = 0; i < text.length; i++) {
      final String t = text[i] ?? '';
      final ComplexField f = cf[i];
      if (f == null || f.instructions == null || f.includeSeparate == null) {
        String closeHyperlink = '';
        String openHyperlink = '';
        if (textStyles[i] != null &&
            textStyles[i].hyperlinkTo != null &&
            textStyles[i].hyperlinkTo.isNotEmpty) {
          _packager.addHyperlink(textStyles[i].hyperlinkTo);
          closeHyperlink = '</w:hyperlink>';
          openHyperlink = '<w:hyperlink r:id="rId${_packager.rIdCount - 1}">';
        }
        d.writeAll(<String>[
          openHyperlink,
          '<w:r>',
          _getTextStyleAsString(
              styleAsHyperlink: textStyles[i] != null &&
                  textStyles[i].hyperlinkTo != null &&
                  textStyles[i].hyperlinkTo.isNotEmpty,
              style: textStyles[i],
              doNotUseGlobalStyle: doNotUseGlobalTextStyle),
          '<w:t xml:space="preserve">$t</w:t></w:r>$multiBreak$closeHyperlink',
        ]);
      } else {
        d.writeAll(<String>[
          '<w:r>',
          _getTextStyleAsString(
              style: textStyles[i],
              doNotUseGlobalStyle: doNotUseGlobalTextStyle),
          '<w:t xml:space="preserve">$t</w:t></w:r><w:r>${_getTextStyleAsString(style: textStyles[i], doNotUseGlobalStyle: doNotUseGlobalTextStyle)}${f.getXml()}</w:r>$multiBreak',
        ]);
      }

      // _addToCharCounters(t);
    }

    final String singleBreak =
        lineOrPageBreak != null && !addBreakAfterEveryItem
            ? lineOrPageBreak.getXml()
            : '';

    d.write('$singleBreak</w:p>');

    return d.toString();
  }

  /// AddAnchorImage inserts an anchor image from a file positioned at the absolute offset coordinates on the page (in EMU), indicated by [horizontalOffsetEMU] and [verticalOffsetEMU]. If both offsets are 0, the anchor image will be placed at the start (left side of the line) of the cursor's current position in the document. Positive offsets push the image to the right and bottom. Negative offsets push the image to the left and top.
  /// Text (if not null) is inserted with the function addMixedText, so [text] is a list of text and is styled with [textStyles].
  /// Given text will behave and wrap as determined by [anchorImageAreaWrap] and [anchorImageTextWrap].
  /// For example, an anchor image can serve as a background image by setting [behindDocument] to true and anchorImageAreaWrap to AnchorImageAreaWrap.wrapNone.
  ///
  /// Width and height are measured in EMU (English Metric Unit) and should be between 1 and 27273042316900.
  /// The width and height should respect the aspect ratio (width / height) of your image to prevent distortion.
  /// You can use convertMillimetersToEMU, convertCentimetersToEMU and convertInchesToEMU to calculate EMU.
  /// The constant emuWidthA4Pct and emuHeightA4Pct values can also be used.
  /// Ensure that the image is compressed to minimize the size of the docx file.
  ///
  /// The image can work like a hyperlink, by setting a URL target in [hyperlinkTo].
  ///
  /// The dimensions of A4:
  ///
  /// width: 210mm/21cm/8.3 inches
  ///
  /// height: 297mm/29.7cm/11.7 inches
  ///
  void addAnchorImage(
    File imageFile,
    int widthEMU,
    int heightEMU,
    int horizontalOffsetEMU,
    int verticalOffsetEMU, {
    int distT = 0,
    int distB = 0,
    int distL = 0,
    int distR = 0,
    bool simplePos = false,
    bool locked = false,
    bool layoutInCell = true,
    bool allowOverlap = true,
    bool behindDocument = false,
    int relativeHeight = 2,
    int simplePosX = 0,
    int simplePosY = 0,
    AnchorImageAreaWrap anchorImageAreaWrap = AnchorImageAreaWrap.wrapSquare,
    AnchorImageTextWrap anchorImageTextWrap = AnchorImageTextWrap.largest,
    HorizontalPositionRelativeBase horizontalPositionRelativeBase =
        HorizontalPositionRelativeBase.column,
    VerticalPositionRelativeBase verticalPositionRelativeBase =
        VerticalPositionRelativeBase.paragraph,
    AnchorImageHorizontalAlignment anchorImageHorizontalAlignment,
    String alternativeTextForImage = '',
    bool noChangeAspect = true,
    bool noChangeArrowheads = true,
    bool noMove = true,
    bool noResize = true,
    bool noSelect = false,
    String hyperlinkTo,
    List<String> text,
    TextStyle imageTextStyle,
    List<TextStyle> textStyles,
    PageStyle pageStyle,
    bool doNotUseGlobalTextStyle = false,
    bool doNotUseGlobalPageStyle = true,
    LineBreak lineOrPageBreak,
    bool addBreakAfterEveryItem = false,
    bool addTab = false,
    int effectExtentL = 0,
    int effectExtentT = 0,
    int effectExtentR = 0,
    int effectExtentB = 0,
    bool flipImageHorizontal = false,
    bool flipImageVertical = false,
    int rotateInEMU = 0,
  }) {
    const int max = 27273042316900;
    final style =
        imageTextStyle ?? TextStyle(textAlignment: TextAlignment.center);
    if (widthEMU < 1 || widthEMU > max || heightEMU < 1 || heightEMU > max) {
      return;
    }

    if (!_bufferClosed) {
      final String path = imageFile.path;
      final String suffix = path.substring(path.lastIndexOf('.') + 1);
      if (mimeTypes.containsKey(suffix)) {
        final bool saved = _packager.addImageFile(imageFile, suffix);
        final String _noChangeAspect = noChangeAspect ? '1' : '0';
        final String _noChangeArrowheads = noChangeArrowheads ? '1' : '0';
        final String _noMove = noMove ? '1' : '0';
        final String _noResize = noResize ? '1' : '0';
        final String _noSelect = noSelect ? '1' : '0';

        final String xfrm = rotateInEMU > 0 ||
                flipImageHorizontal ||
                flipImageVertical
            ? '<a:xfrm rot="$rotateInEMU" flipH="$flipImageHorizontal" flipV="$flipImageVertical"><a:off x="0" y="0"/><a:ext cx="$widthEMU" cy="$heightEMU"/></a:xfrm>'
            : '';

        final int mediaIdCount = _packager.rIdCount - 1;

        String hyperlink = '';
        if (hyperlinkTo != null && hyperlinkTo.isNotEmpty) {
          hyperlink =
              '<a:hlinkClick xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:id="rId${_packager.rIdCount}"/>';
          _packager.addHyperlink(hyperlinkTo);
        }

        const String openParagraph = '<w:p>';
        const String closeParagraph = '</w:p>';
        final String openPpr =
            '${_getParagraphStyleAsString(textStyle: style, doNotUseGlobalStyle: doNotUseGlobalTextStyle)}';
        const String openR = '<w:r>';
        const String closeR = '</w:r>';

        if (saved) {
          _docx.write(
              '$openParagraph$openPpr$openR<w:drawing><wp:anchor behindDoc="$behindDocument" distT="$distT" distB="$distB" distL="$distL" distR="$distR" simplePos="$simplePos" locked="$locked" layoutInCell="$layoutInCell" allowOverlap="$allowOverlap" relativeHeight="$relativeHeight"><wp:simplePos x="$simplePosX" y="$simplePosY" /><wp:positionH relativeFrom="${getValueFromEnum(horizontalPositionRelativeBase)}"><wp:posOffset>$horizontalOffsetEMU</wp:posOffset></wp:positionH><wp:positionV relativeFrom="${getValueFromEnum(verticalPositionRelativeBase)}"><wp:posOffset>$verticalOffsetEMU</wp:posOffset></wp:positionV><wp:extent cx="$widthEMU" cy="$heightEMU"/><wp:effectExtent l="$effectExtentL" t="$effectExtentT" r="$effectExtentR" b="$effectExtentB"/><wp:${getValueFromEnum(anchorImageAreaWrap)} wrapText="${getValueFromEnum(anchorImageTextWrap)}"/><wp:docPr id="$mediaIdCount" name="Image$mediaIdCount" descr="$alternativeTextForImage">$hyperlink</wp:docPr><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="$_noChangeAspect" noMove="$_noMove" noResize="$_noResize" noSelect="$_noSelect"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="$mediaIdCount" name="Image$mediaIdCount" descr="$alternativeTextForImage"></pic:cNvPr><pic:cNvPicPr><a:picLocks noChangeAspect="$_noChangeAspect" noMove="$_noMove" noResize="$_noResize" noSelect="$_noSelect" noChangeArrowheads="$_noChangeArrowheads"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId$mediaIdCount"></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr bwMode="auto">$xfrm</pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>$closeR$closeParagraph');
          addMixedText(
            text,
            textStyles,
            pageStyle: pageStyle,
            doNotUseGlobalTextStyle: doNotUseGlobalTextStyle,
            doNotUseGlobalPageStyle: doNotUseGlobalPageStyle,
            lineOrPageBreak: lineOrPageBreak,
            addBreakAfterEveryItem: addBreakAfterEveryItem,
            addTab: addTab,
          );
        }
      }
    }
  }

  /// AddImage inserts an inline image from a file at the current position in the buffer. AddImage is useful for stand-alone images that do not have accompanying text.
  /// If you want multiple images in the same paragraph, use AddImages. If you need an anchor image or an image surrounded with text, use addAnchorImage. You can also use tables to create a grid of images and text as an alternative for anchor images.
  ///
  /// Ensure that the image is compressed to minimize the size of the docx file.
  ///
  /// Width and height are measured in EMU (English Metric Unit) and should be between 1 and 27273042316900.
  /// The width and height should respect the aspect ratio (width / height) of your image to prevent distortion.
  /// You can use convertMillimetersToEMU, convertCentimetersToEMU and convertInchesToEMU to calculate EMU.
  /// The constant emuWidthA4Pct and emuHeightA4Pct values can also be used.
  ///
  /// The image can work like a hyperlink, by setting a URL target in [hyperlinkTo].
  ///
  /// Text styling that affects the paragraph's appearance also affects the appearance of the image, for example: textAlignment, verticalTextAlignment, paragraphShading, paragraphBorders and paragraphBorderOnAllSides.
  ///
  /// The dimensions of A4:
  ///
  /// width: 210mm/21cm/8.3 inches
  ///
  /// height: 297mm/29.7cm/11.7 inches
  ///
  void addImage(
    File imageFile,
    int widthEMU,
    int heightEMU, {
    String alternativeTextForImage = '',
    bool noChangeAspect = true,
    bool noMove = true,
    bool noResize = true,
    bool noSelect = false,
    TextStyle textStyle,
    bool doNotUseGlobalStyle = true,
    String hyperlinkTo,
    int effectExtentL = 0,
    int effectExtentT = 0,
    int effectExtentR = 0,
    int effectExtentB = 0,
    bool flipImageHorizontal = false,
    bool flipImageVertical = false,
    int rotateInEMU = 0,
  }) =>
      _insertInlineImage(
        imageFile,
        widthEMU,
        heightEMU,
        alternativeTextForImage: alternativeTextForImage,
        noChangeAspect: noChangeAspect,
        noMove: noMove,
        noResize: noResize,
        noSelect: noSelect,
        textStyle: textStyle,
        doNotUseGlobalStyle: doNotUseGlobalStyle,
        encloseInParagraph: true,
        hyperlinkTo: hyperlinkTo,
        flipImageHorizontal: flipImageHorizontal,
        flipImageVertical: flipImageVertical,
        rotateInEMU: rotateInEMU,
      );

  /// Add multiple inline images uniformly inside the same paragraph.
  /// Width and height are measured in EMU (English Metric Unit) and should be between 1 and 27273042316900.
  /// All given images will be placed with the same widthEMU and heightEMU.
  /// Ensure that all images are compressed to minimize the size of the docx file.
  ///
  /// The width and height should respect the aspect ratio (width / height) of your images to prevent distortion.
  /// You can use convertMillimetersToEMU, convertCentimetersToEMU and convertInchesToEMU to calculate EMU.
  /// The constant emuWidthA4Pct and emuHeightA4Pct values can also be used.
  ///
  /// If given, the [hyperlinksTo] list for the images should have the same length as imageFiles.
  /// You can add distance between the images with [spaces] which mimic spacebar presses.
  void addImages(
    List<File> imageFiles,
    int widthEMU,
    int heightEMU, {
    List<String> alternativeTextListForImage,
    List<String> hyperlinksTo,
    bool noChangeAspect = true,
    bool noMove = true,
    bool noResize = true,
    bool noSelect = false,
    TextStyle textStyle,
    bool doNotUseGlobalStyle = true,
    int spaces = 0,
    int effectExtentL = 0,
    int effectExtentT = 0,
    int effectExtentR = 0,
    int effectExtentB = 0,
    bool flipImageHorizontal = false,
    bool flipImageVertical = false,
    int rotateInEMU = 0,
  }) {
    final style = textStyle ?? TextStyle(textAlignment: TextAlignment.center);

    final List<String> desc = alternativeTextListForImage != null &&
            alternativeTextListForImage.length == imageFiles.length
        ? alternativeTextListForImage
        : List.generate(imageFiles.length, (index) => '');

    final List<String> hyperlinks =
        hyperlinksTo != null && hyperlinksTo.length == imageFiles.length
            ? hyperlinksTo
            : List.generate(imageFiles.length, (index) => '');

    final String openParagraphWithPpr =
        '<w:p>${_getParagraphStyleAsString(textStyle: style, doNotUseGlobalStyle: doNotUseGlobalStyle)}<w:r>';
    _docx.write(openParagraphWithPpr);
    final String addSpaces = spaces > 0
        ? '<w:t xml:space="preserve">${List<String>.generate(spaces, (index) => ' ').join()}</w:t>'
        : '';
    for (int i = 0; i < imageFiles.length; i++) {
      _insertInlineImage(
        imageFiles[i],
        widthEMU,
        heightEMU,
        alternativeTextForImage: desc[i] ?? '',
        noChangeAspect: noChangeAspect,
        noMove: noMove,
        noResize: noResize,
        noSelect: noSelect,
        textStyle: textStyle,
        doNotUseGlobalStyle: doNotUseGlobalStyle,
        encloseInParagraph: false,
        hyperlinkTo: hyperlinks[i] ?? '',
        effectExtentL: effectExtentL,
        effectExtentT: effectExtentT,
        effectExtentR: effectExtentR,
        effectExtentB: effectExtentB,
        flipImageHorizontal: flipImageHorizontal,
        flipImageVertical: flipImageVertical,
        rotateInEMU: rotateInEMU,
      );
      if (addSpaces.isNotEmpty) {
        _docx.write(addSpaces);
      }
    }
    _docx.write('</w:r></w:p>');
  }

  /// Create an inline image with a caption. The function itself is just syntactic sugar for creating a table with 1 column and 2 rows followed by placing an inline image in one row and text in the other.
  ///
  /// [tableWidthInTwips] determines the width of the table (measured in twips).
  /// [captionAppearsBelowImage] determines if the caption is displayed above (false) or below (true; default) the image.
  /// [hyperlinkTo] applies to the image, not the caption. If the caption needs to be a hyperlink, you can use TextStyle.hyperlinkTo or add the complexField(instructions: 'HYPERLINK').
  void addImageWithCaption(
    File imageFile,
    int widthEMU,
    int heightEMU, {
    int tableWidthInTwips = 8000,
    bool captionAppearsBelowImage = true,
    Shading shadingCaption,
    Shading shadingImage,
    List<String> text,
    List<TextStyle> textStyles,
    String alternativeTextForImage = '',
    bool noChangeAspect = true,
    bool noMove = true,
    bool noResize = true,
    bool noSelect = false,
    TextStyle textStyle,
    String hyperlinkTo,
    int effectExtentL = 0,
    int effectExtentT = 0,
    int effectExtentR = 0,
    int effectExtentB = 0,
    bool flipImageHorizontal = false,
    bool flipImageVertical = false,
    int rotateInEMU = 0,
    List<TableBorder> tableBorders,
    TableTextAlignment tableTextAlignment,
    TableCellVerticalAlignment tableCellVerticalAlignment,
    bool doNotUseGlobalTextStyle = false,
    bool doNotUseGlobalPageStyle = true,
    bool addTab = false,
    List<ComplexField> complexFields,
  }) {
    if (!_bufferClosed) {
      final xmlImage = _getCachedInlineImage(
        imageFile,
        widthEMU,
        heightEMU,
        alternativeTextForImage: alternativeTextForImage,
        noChangeAspect: noChangeAspect,
        noMove: noMove,
        noResize: noResize,
        noSelect: noSelect,
        effectExtentB: effectExtentB,
        effectExtentL: effectExtentL,
        effectExtentR: effectExtentR,
        effectExtentT: effectExtentT,
        encloseInParagraph: true,
        flipImageHorizontal: flipImageHorizontal,
        flipImageVertical: flipImageVertical,
        rotateInEMU: rotateInEMU,
        hyperlinkTo: hyperlinkTo,
      );
      final xmlCaption = _getCachedAddMixedText(
        text,
        textStyles,
        complexFields: complexFields,
        doNotUseGlobalPageStyle: doNotUseGlobalPageStyle,
        doNotUseGlobalTextStyle: doNotUseGlobalTextStyle,
        addTab: addTab,
      );

      final TableRow topRow = TableRow(
        tableRowProperties: TableRowProperties(
            tableCellSpacingWidthType: PreferredWidthType.auto),
        tableCells: [
          TableCell(
            tableCellProperties: TableCellProperties(
              preferredWidthType: PreferredWidthType.auto,
              verticalAlignment: tableCellVerticalAlignment,
              shading: captionAppearsBelowImage ? shadingImage : shadingCaption,
            ),
            xmlContent: captionAppearsBelowImage ? xmlImage : xmlCaption,
          )
        ],
      );

      final TableRow bottomRow = TableRow(
        tableRowProperties: TableRowProperties(
          tableCellSpacingWidthType: PreferredWidthType.auto,
        ),
        tableCells: [
          TableCell(
            tableCellProperties: TableCellProperties(
              preferredWidthType: PreferredWidthType.auto,
              verticalAlignment: tableCellVerticalAlignment,
              shading:
                  !captionAppearsBelowImage ? shadingImage : shadingCaption,
            ),
            xmlContent: !captionAppearsBelowImage ? xmlImage : xmlCaption,
          )
        ],
      );

      final Table t = Table(
        gridColumnWidths: [tableWidthInTwips ?? 8000],
        tableProperties: TableProperties(
          tableBorders: tableBorders,
          tableTextAlignment: tableTextAlignment,
        ),
        tableRows: [topRow, bottomRow],
      );

      attachTable(t);
    }
  }

  void insertImageInTableCell({
    @required TableCell tableCell,
    @required File imageFile,
    @required int widthEMU,
    @required int heightEMU,
    String alternativeTextForImage = '',
    bool noChangeAspect = true,
    bool noMove = true,
    bool noResize = true,
    bool noSelect = false,
    TextStyle textStyle,
    bool doNotUseGlobalStyle = true,
    String hyperlinkTo,
    int effectExtentL = 0,
    int effectExtentT = 0,
    int effectExtentR = 0,
    int effectExtentB = 0,
    bool flipImageHorizontal = false,
    bool flipImageVertical = false,
    int rotateInEMU = 0,
  }) {
    const int max = 27273042316900;
    if (widthEMU < 1 || widthEMU > max || heightEMU < 1 || heightEMU > max) {
      return;
    }

    if (!_bufferClosed) {
      tableCell.xmlContent = _getCachedInlineImage(
        imageFile,
        widthEMU,
        heightEMU,
        alternativeTextForImage: alternativeTextForImage,
        noChangeAspect: noChangeAspect,
        noMove: noMove,
        noResize: noResize,
        noSelect: noSelect,
        textStyle: textStyle,
        doNotUseGlobalStyle: doNotUseGlobalStyle,
        encloseInParagraph: true,
        hyperlinkTo: hyperlinkTo,
        effectExtentL: effectExtentL,
        effectExtentT: effectExtentT,
        effectExtentR: effectExtentR,
        effectExtentB: effectExtentB,
        flipImageHorizontal: flipImageHorizontal,
        flipImageVertical: flipImageVertical,
        rotateInEMU: rotateInEMU,
      );
    }
  }

  void insertImagesInTableCell({
    @required TableCell tableCell,
    @required List<File> imageFiles,
    @required int widthEMU,
    @required int heightEMU,
    List<String> alternativeTextListForImage,
    List<String> hyperlinksTo,
    bool noChangeAspect = true,
    bool noMove = true,
    bool noResize = true,
    bool noSelect = false,
    TextStyle textStyle,
    bool doNotUseGlobalStyle = true,
    int spaces = 0,
    int effectExtentL = 0,
    int effectExtentT = 0,
    int effectExtentR = 0,
    int effectExtentB = 0,
    bool flipImageHorizontal = false,
    bool flipImageVertical = false,
    int rotateInEMU = 0,
  }) {
    final StringBuffer s = StringBuffer();
    final style = textStyle ?? TextStyle(textAlignment: TextAlignment.center);

    final List<String> desc = alternativeTextListForImage != null &&
            alternativeTextListForImage.length == imageFiles.length
        ? alternativeTextListForImage
        : List.generate(imageFiles.length, (index) => '');

    final List<String> hyperlinks =
        hyperlinksTo != null && hyperlinksTo.length == imageFiles.length
            ? hyperlinksTo
            : List.generate(imageFiles.length, (index) => '');

    final String openParagraphWithPpr =
        '<w:p>${_getParagraphStyleAsString(textStyle: style, doNotUseGlobalStyle: doNotUseGlobalStyle)}<w:r>';
    s.write(openParagraphWithPpr);
    final String addSpaces = spaces > 0
        ? '<w:t xml:space="preserve">${List<String>.generate(spaces, (index) => ' ').join()}</w:t>'
        : '';
    for (int i = 0; i < imageFiles.length; i++) {
      s.write(_getCachedInlineImage(
        imageFiles[i],
        widthEMU,
        heightEMU,
        alternativeTextForImage: desc[i] ?? '',
        noChangeAspect: noChangeAspect,
        noMove: noMove,
        noResize: noResize,
        noSelect: noSelect,
        textStyle: textStyle,
        doNotUseGlobalStyle: doNotUseGlobalStyle,
        encloseInParagraph: false,
        hyperlinkTo: hyperlinks[i] ?? '',
        effectExtentL: effectExtentL,
        effectExtentT: effectExtentT,
        effectExtentR: effectExtentR,
        effectExtentB: effectExtentB,
        flipImageHorizontal: flipImageHorizontal,
        flipImageVertical: flipImageVertical,
        rotateInEMU: rotateInEMU,
      ));
      if (addSpaces.isNotEmpty) {
        s.write(addSpaces);
      }
    }
    s.write('</w:r></w:p>');
    tableCell.xmlContent = s.toString();
  }

  void _insertInlineImage(
    File imageFile,
    int widthEMU,
    int heightEMU, {
    String alternativeTextForImage = '',
    bool noChangeAspect = true,
    bool noMove = true,
    bool noResize = true,
    bool noSelect = false,
    TextStyle textStyle,
    bool doNotUseGlobalStyle = true,
    bool encloseInParagraph = true,
    String hyperlinkTo,
    int effectExtentL = 0,
    int effectExtentT = 0,
    int effectExtentR = 0,
    int effectExtentB = 0,
    bool flipImageHorizontal = false,
    bool flipImageVertical = false,
    int rotateInEMU = 0,
  }) {
    const int max = 27273042316900;
    if (widthEMU < 1 || widthEMU > max || heightEMU < 1 || heightEMU > max) {
      return;
    }

    if (!_bufferClosed) {
      _docx.write(_getCachedInlineImage(
        imageFile,
        widthEMU,
        heightEMU,
        alternativeTextForImage: alternativeTextForImage,
        noChangeAspect: noChangeAspect,
        noMove: noMove,
        noResize: noResize,
        noSelect: noSelect,
        textStyle: textStyle,
        doNotUseGlobalStyle: doNotUseGlobalStyle,
        encloseInParagraph: encloseInParagraph,
        hyperlinkTo: hyperlinkTo,
        effectExtentL: effectExtentL,
        effectExtentT: effectExtentT,
        effectExtentR: effectExtentR,
        effectExtentB: effectExtentB,
        flipImageHorizontal: flipImageHorizontal,
        flipImageVertical: flipImageVertical,
        rotateInEMU: rotateInEMU,
      ));
    }
  }

  String _getCachedInlineImage(File imageFile, int widthEMU, int heightEMU,
      {String alternativeTextForImage = '',
      bool noChangeAspect = true,
      bool noMove = true,
      bool noResize = true,
      bool noSelect = false,
      TextStyle textStyle,
      bool doNotUseGlobalStyle = true,
      bool encloseInParagraph = true,
      String hyperlinkTo,
      int effectExtentL = 0,
      int effectExtentT = 0,
      int effectExtentR = 0,
      int effectExtentB = 0,
      bool flipImageHorizontal = false,
      bool flipImageVertical = false,
      int rotateInEMU = 0}) {
    final style = textStyle ?? TextStyle(textAlignment: TextAlignment.center);
    final StringBuffer d = StringBuffer();
    final String path = imageFile.path;
    final String suffix = path.substring(path.lastIndexOf('.') + 1);
    if (mimeTypes.containsKey(suffix)) {
      final bool saved = _packager.addImageFile(imageFile, suffix);
      if (saved) {
        final String _noChangeAspect = noChangeAspect ? '1' : '0';
        final String _noMove = noMove ? '1' : '0';
        final String _noResize = noResize ? '1' : '0';
        final String _noSelect = noSelect ? '1' : '0';

        final String xfrm = rotateInEMU > 0 ||
                flipImageHorizontal ||
                flipImageVertical
            ? '<a:xfrm rot="$rotateInEMU" flipH="$flipImageHorizontal" flipV="$flipImageVertical"><a:off x="0" y="0"/><a:ext cx="$widthEMU" cy="$heightEMU"/></a:xfrm>'
            : '';

        final int mediaIdCount = _packager.rIdCount - 1;

        String hyperlink = '';
        if (hyperlinkTo != null && hyperlinkTo.isNotEmpty) {
          hyperlink =
              '<a:hlinkClick xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:id="rId${_packager.rIdCount}"/>';
          _packager.addHyperlink(hyperlinkTo);
        }

        final String openParagraph = encloseInParagraph ? '<w:p>' : '';
        final String closeParagraph = encloseInParagraph ? '</w:p>' : '';
        final String openPpr = encloseInParagraph
            ? '${_getParagraphStyleAsString(textStyle: style, doNotUseGlobalStyle: doNotUseGlobalStyle)}'
            : '';
        final String openR = encloseInParagraph ? '<w:r>' : '';
        final String closeR = encloseInParagraph ? '</w:r>' : '';

        d.write(
            '$openParagraph$openPpr$openR<w:drawing><wp:inline><wp:extent cx="$widthEMU" cy="$heightEMU"/><wp:effectExtent l="$effectExtentL" t="$effectExtentT" r="$effectExtentR" b="$effectExtentB"/><wp:docPr id="$mediaIdCount" name="Image$mediaIdCount" descr="$alternativeTextForImage">$hyperlink</wp:docPr><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="$_noChangeAspect" noMove="$_noMove" noResize="$_noResize" noSelect="$_noSelect"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="$mediaIdCount" name="Image$mediaIdCount" descr="$alternativeTextForImage"></pic:cNvPr><pic:cNvPicPr><a:picLocks noChangeAspect="$_noChangeAspect" noMove="$_noMove" noResize="$_noResize" noSelect="$_noSelect" noChangeArrowheads="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId$mediaIdCount"></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr bwMode="auto">$xfrm</pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>$closeR$closeParagraph');
      }
    }
    return d.toString();
  }

  /// Create the .docx file with the content stored in the buffer of DocXBuilder. This also closes the buffer.
  ///
  /// If provided, [documentTitle], [documentSubject], [documentDescription] and [documentCreator] can be registered in the document.
  ///
  /// After processing the .docx file, do not forget to call clear to free resources and reopen the buffer.
  Future<File> createDocXFile({
    String documentTitle = '',
    String documentSubject = '',
    String documentDescription = '',
    String documentCreator = '',
  }) async {
    try {
      if (!_bufferClosed) {
        _docx.write(_getPageStyleAsString(style: _globalPageStyle));
        _docx.write('</w:body></w:document>');
        _bufferClosed = true;
      }

      final File f = await _packager.bundle(
        _docx.toString(),
        // chars: _charCount,
        // charsWithSpaces: _charCountWithSpaces,
        // paragraphs: _parCount,
        documentTitle: documentTitle ?? '',
        documentSubject: documentSubject ?? '',
        documentDescription: documentDescription ?? '',
        documentCreator: documentCreator ?? '',
      );
      return f;
    } catch (e) {
      throw StateError('DOCX BUILDER ERROR: $e');
    }
  }

  /// Deletes all data and cache, so a new document can be created (or when DocXBuilder is no longer needed).
  /// This function also destroys the source file produced by createDocxFile so be sure to save or process the source file before calling clear.
  void clear({bool resetDocX = true}) {
    _docx.clear();
    _documentBackgroundColor = null;
    _globalPageStyle = null;
    _globalTextStyle = null;
    _firstPageHeader = null;
    _oddPageHeader = null;
    _evenPageHeader = null;
    _firstPageFooter = null;
    _oddPageFooter = null;
    _evenPageFooter = null;
    _insertHeadersAndFootersInThisSection = false;
    // _charCount = 0;
    // _charCountWithSpaces = 0;
    // _parCount = 0;
    _headerCounter = 1;
    _footerCounter = 1;
    _packager.destroyCache();
    if (resetDocX) {
      _initDocX();
    }
  }
}
