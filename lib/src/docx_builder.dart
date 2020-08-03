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

export 'styles/style_classes/index.dart';
export 'styles/style_containers/index.dart';
export 'styles/style_enums.dart';

/// DocXBuilder is used to create docx files by using OOXML.
/// Global styling is set with setDocumentBackgroundColor, setGlobalTextStyle and setGlobalPageStyle.
///
/// Write text to the document with addText or addMixedText. These functions add a paragraph of text with both visual (e.g. font size, color, font weight) and complex styling (e.g. numbered list, text frame, complex fields).
///
/// Creating a table is done in 4 steps: make table cells that have content like images and text; make table rows that have table cells; make a table that has table rows; when a table is complete, attach the table to the document.
/// Except for attachTable, all table functions start with "insert" (e.g. insertTextInTableCell).
///
/// OOXML defines 2 types of images: inline and anchor.
/// An inline image is placed at the cursor's current position as if it was a large area of text. An anchor image is a floating image and can be placed anywhere you like on the page.
/// Except for addAnchorImage, all image functions involve inline images (e.g. addImage, addImages, insertImageInTableCell).
///
/// CreateDocXFile is used to create the .docx file.
/// After creating the docx file, clear() is used to destroy the cache.
class DocXBuilder {
  _p.Packager _packager;
  final StringBuffer _docx = StringBuffer();

  String _debugString;

  /// OOXML does not allow the last paragraph in the document to have its own section properties; instead, they should be placed directly in the body of the document.
  /// To prevent problems with this rule, docxBuilder adds an empty paragraph (i.e. a last paragraph without section properties) at the end of the document.
  ///
  /// If you know the last paragraph does not have its own section rules and you do not want an empty line at the end of the document, you can set this value to false.
  bool addEmptyParagraphAtEndOfDocument = true;

  /// If bookmarks were added anywhere in the document (e.g. addText('Bookmarked!', setBookmarkName: 'bookmark1')), they will be added here.
  Set<String> bookmarks = <String>{};

  /// If comments were added anywhere in the document with the function attachComment, they will be added here.
  /// If this map is not empty, when createDocx is called, the packager will create and include the file "comments.xml" and all necessary references.
  Map<int, String> comments = <int, String>{};

  /// The hyperlink styling is by default equivalent to TextStyle(color: '000080', underline: Underline.single), but you can set another textStyle to change it.
  TextStyle hyperlinkTextStyle;

  bool _bufferClosed = false;
  // int _charCount = 0;
  // int _charCountWithSpaces = 0;
  // int _parCount = 0;
  int _headerCounter = 1;
  int _footerCounter = 1;
  int _anchorCounter = 0;
  int _commentCounter = 0;

  String _documentBackgroundColor;
  String get documentBackgroundColor => _documentBackgroundColor;

  /// The global style for text. Except for HighlightColor, colors are in 'RRGGBB' format.
  /// The global text style can be overridden by giving addText and addMixedText a textStyle.
  /// The hyperlinkTo property of the global text style is always ignored. Instead, hyperlinks should be added manually with addText or addMixedText.
  TextStyle globalTextStyle = TextStyle();

  /// The global page style, such as orientation, size, and page borders.
  /// This page style determines the content of the final Sections Properties inside the body tag of the document.
  PageStyle globalPageStyle = PageStyle().getDefaultPageStyle();

  Header _firstPageHeader;
  Header _oddPageHeader;
  Header _evenPageHeader;
  Footer _firstPageFooter;
  Footer _oddPageFooter;
  Footer _evenPageFooter;
  bool _insertHeadersAndFootersInThisSection = false;
  String _customNumberingXml;
  bool _includeNumberingXml = false;

  final String mimetype =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

  final int emuWidthA4 = 7560000;
  final int emuHeightA4 = 10692000;
  final int twipsWidthA4 = 11906;
  final int twipsHeightA4 = 16838;

  final String bullet = '•';
  final String bulletArrow = '‣';
  final String bulletDiamond = '⬩';
  final String bulletHyphen = '⁃';
  final String bulletSquare = '▪';
  final String bulletSquareWhite = '▫';
  final String bulletWhite = '◦';

  /// Convert EMUs to Twips (twentieth of point).
  int convertEMUToTwips(int emu) => (emu / 635).round();

  /// Convert Twips to EMUS.
  int convertTwipsToEMU(int twips) => twips * 635;

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

  // Blank page can also be added with sectPr, sectType.nextPage and then create a new section which inserts a new page
  // void addBlankPageA4() => _docx.write(_c.blankPageA4);

  /// A custom numbering list can be set once for the entire document.
  ///
  /// Maximum level of depth is 8.
  void setCustomNumberingList({
    @required List<NumberFormat> numberFormatForEachLevel,
    @required List<String> charactersForEachLevel,
    int tabSpace = 420,
    bool justifyToLeft = true,
    NumberingSuffix suffix,
    bool isLgl = false,
  }) {
    if (_customNumberingXml == null &&
        numberFormatForEachLevel != null &&
        charactersForEachLevel != null) {
      final List<NumberFormat> formats = <NumberFormat>[
        ...numberFormatForEachLevel,
        ...List<NumberFormat>.generate(8, (index) => null),
      ].sublist(0, 8);

      final List<String> chars = <String>[
        ...charactersForEachLevel,
        ...List<String>.generate(8, (index) => null),
      ].sublist(0, 8);

      final String jc = justifyToLeft ? 'left' : 'right';

      final String suff =
          suffix != null ? '<w:suff w:val="${getValueFromEnum(suffix)}"/>' : '';

      final String lgl = isLgl ?? false ? '<w:isLgl/>' : '';

      final StringBuffer c = StringBuffer()
        ..write(
            '<w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="multilevel"/>');
      for (int i = 0; i < formats.length; i++) {
        final String ppr =
            '<w:pPr><w:tabs><w:tab w:val="left" w:pos="${tabSpace * (i + 1)}"/></w:tabs><w:ind w:left="${tabSpace * (i + 1)}" w:leftChars="0" w:hanging="$tabSpace" w:firstLineChars="0"/></w:pPr>';
        const String rpr = '<w:rPr><w:rFonts w:hint="default"/></w:rPr>';
        if (formats[i] != null) {
          c.write(
              '<w:lvl w:ilvl="$i" w:tentative="0"><w:start w:val="1"/>$suff$lgl<w:numFt w:val="${getValueFromEnum(formats[i])}"/><w:lvlText w:val="${chars[i] ?? "%${i + 1}"}"/><w:lvlJc w:val="$jc"/>$ppr$rpr</w:lvl>');
        }
      }
      c.write('</w:abstractNum>');
      _customNumberingXml = c.toString();
    }
  }

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
          bool styleAsHyperlink = false;
          String closeHyperlink = '';
          String openHyperlink = '';
          if (textStyles[i] != null &&
              textStyles[i].hyperlinkTo != null &&
              textStyles[i].hyperlinkTo.isNotEmpty) {
            final tags =
                _getOpenAndCloseHyperlinkTags(textStyles[i].hyperlinkTo);
            openHyperlink = tags[0];
            closeHyperlink = tags[1];
            styleAsHyperlink = true;
          }
          b.writeAll(<String>[
            openHyperlink,
            '<w:r>',
            _getTextStyleAsString(
                styleAsHyperlink: styleAsHyperlink,
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
  /// If no style is given, then the globalPageStyle is used.
  String _getPageStyleAsString(
      {PageStyle style, bool doNotUseGlobalStyle = false}) {
    final PageStyle pageStyle = style ?? PageStyle();
    String spr = _b.SectPr.getSpr(
      cols: doNotUseGlobalStyle
          ? pageStyle.cols
          : pageStyle.cols ?? globalPageStyle.cols,
      colsHaveSeparator: doNotUseGlobalStyle
          ? pageStyle.colsHaveSeparator
          : pageStyle.colsHaveSeparator ?? globalPageStyle.colsHaveSeparator,
      colsHaveEqualWidth: doNotUseGlobalStyle
          ? pageStyle.colsHaveEqualWidth
          : pageStyle.colsHaveEqualWidth ?? globalPageStyle.colsHaveEqualWidth,
      sectType: doNotUseGlobalStyle
          ? pageStyle.sectType
          : pageStyle.sectType ?? globalPageStyle.sectType,
      pageSzHeight: doNotUseGlobalStyle
          ? pageStyle.pageSzHeight
          : pageStyle.pageSzHeight ?? globalPageStyle.pageSzHeight,
      pageSzWidth: doNotUseGlobalStyle
          ? pageStyle.pageSzWidth
          : pageStyle.pageSzWidth ?? globalPageStyle.pageSzWidth,
      pageOrientation: doNotUseGlobalStyle
          ? pageStyle.pageOrientation
          : pageStyle.pageOrientation ?? globalPageStyle.pageOrientation,
      sectVerticalAlign: doNotUseGlobalStyle
          ? pageStyle.sectVerticalAlign
          : pageStyle.sectVerticalAlign ?? globalPageStyle.sectVerticalAlign,
      pageBorders: doNotUseGlobalStyle
          ? pageStyle.pageBorders
          : pageStyle.pageBorders ?? globalPageStyle.pageBorders,
      pageBorderDisplay: doNotUseGlobalStyle
          ? pageStyle.pageBorderDisplay
          : pageStyle.pageBorderDisplay ?? globalPageStyle.pageBorderDisplay,
      pageBorderOffsetBasedOnText: doNotUseGlobalStyle
          ? pageStyle.pageBorderOffsetBasedOnText
          : pageStyle.pageBorderOffsetBasedOnText ??
              globalPageStyle.pageBorderOffsetBasedOnText,
      pageBorderIsRenderedAboveText: doNotUseGlobalStyle
          ? pageStyle.pageBorderIsRenderedAboveText
          : pageStyle.pageBorderIsRenderedAboveText ??
              globalPageStyle.pageBorderIsRenderedAboveText,
      pageNumberingFormat: doNotUseGlobalStyle
          ? pageStyle.pageNumberingFormat
          : pageStyle.pageNumberingFormat ??
              globalPageStyle.pageNumberingFormat,
      pageNumberingStart: doNotUseGlobalStyle
          ? pageStyle.pageNumberingStart
          : pageStyle.pageNumberingStart ?? globalPageStyle.pageNumberingStart,
      lineNumbering: doNotUseGlobalStyle
          ? pageStyle.lineNumbering
          : pageStyle.lineNumbering ?? globalPageStyle.lineNumbering,
      pageMargin: doNotUseGlobalStyle
          ? pageStyle.pageMargin
          : pageStyle.pageMargin ?? globalPageStyle.pageMargin,
    );
    if (_insertHeadersAndFootersInThisSection) {
      spr = _addHeadersAndFootersToSectPr(spr);
      _insertHeadersAndFootersInThisSection = false;
    }
    return spr;
  }

  /// Obtain the XML string of the paragraph style, such as text alignment.
  /// If no style is given, then the globalTextStyle is used (unless [doNotUseGlobalStyle] is set to true).
  String _getParagraphStyleAsString({
    TextStyle textStyle,
    bool doNotUseGlobalStyle = false,
  }) {
    final TextStyle style = textStyle ?? TextStyle();
    if (style?.numberingList != null ?? false) {
      _includeNumberingXml = true;
    }
    return _b.Ppr.getPpr(
      textFrame: textStyle?.textFrame != null ? textStyle.textFrame : null,
      keepLines: doNotUseGlobalStyle
          ? style.keepLines
          : style.keepLines ?? globalTextStyle.keepLines,
      keepNext: doNotUseGlobalStyle
          ? style.keepNext
          : style.keepNext ?? globalTextStyle.keepNext,
      paragraphBorderOnAllSides: doNotUseGlobalStyle
          ? style.paragraphBorderOnAllSides
          : style.paragraphBorderOnAllSides ??
              globalTextStyle.paragraphBorderOnAllSides,
      paragraphBorders: doNotUseGlobalStyle
          ? style.paragraphBorders
          : style.paragraphBorders ?? globalTextStyle.paragraphBorders,
      paragraphIndent: doNotUseGlobalStyle
          ? style.paragraphIndent
          : style.paragraphIndent ?? globalTextStyle.paragraphIndent,
      paragraphShading: doNotUseGlobalStyle
          ? style.paragraphShading
          : style.paragraphShading ?? globalTextStyle.paragraphShading,
      spacing: doNotUseGlobalStyle
          ? style.paragraphSpacing
          : style.paragraphSpacing ?? globalTextStyle.paragraphSpacing,
      tabs:
          doNotUseGlobalStyle ? style.tabs : style.tabs ?? globalTextStyle.tabs,
      textAlignment: doNotUseGlobalStyle
          ? style.textAlignment
          : style.textAlignment ?? globalTextStyle.textAlignment,
      verticalTextAlignment: doNotUseGlobalStyle
          ? style.verticalTextAlignment
          : style.verticalTextAlignment ??
              globalTextStyle.verticalTextAlignment,
      numberingList: doNotUseGlobalStyle
          ? style.numberingList
          : style.numberingList ?? globalTextStyle.numberingList,
      numberLevelInList: doNotUseGlobalStyle
          ? style.numberLevelInList
          : style.numberLevelInList ?? globalTextStyle.numberLevelInList,
    );
  }

  /// Obtain the XML string of the text style.
  /// If no style is given, then the globalTextStyle is used (unless [doNotUseGlobalStyle] is set to true).
  String _getTextStyleAsString(
      {TextStyle style,
      bool doNotUseGlobalStyle = false,
      bool styleAsHyperlink = false}) {
    final TextStyle textStyle = style ?? TextStyle();
    return _b.Rpr.getRpr(
      styleAsHyperlink: styleAsHyperlink,
      bold: doNotUseGlobalStyle
          ? textStyle.bold
          : textStyle.bold ?? globalTextStyle.bold,
      caps: doNotUseGlobalStyle
          ? textStyle.caps
          : textStyle.caps ?? globalTextStyle.caps,
      doubleStrike: doNotUseGlobalStyle
          ? textStyle.doubleStrike
          : textStyle.doubleStrike ?? globalTextStyle.doubleStrike,
      fontColor: doNotUseGlobalStyle
          ? textStyle.fontColor
          : textStyle.fontColor ?? globalTextStyle.fontColor,
      fontName: doNotUseGlobalStyle
          ? textStyle.fontName
          : textStyle.fontName ?? globalTextStyle.fontName,
      fontNameComplexScript: doNotUseGlobalStyle
          ? textStyle.fontNameComplexScript
          : textStyle.fontNameComplexScript ??
              globalTextStyle.fontNameComplexScript,
      fontSize: doNotUseGlobalStyle
          ? textStyle.fontSize
          : textStyle.fontSize ?? globalTextStyle.fontSize,
      highlightColor: doNotUseGlobalStyle
          ? textStyle.highlightColor
          : textStyle.highlightColor ?? globalTextStyle.highlightColor,
      italic: doNotUseGlobalStyle
          ? textStyle.italic
          : textStyle.italic ?? globalTextStyle.italic,
      rtlText: doNotUseGlobalStyle
          ? textStyle.rtlText
          : textStyle.rtlText ?? globalTextStyle.rtlText,
      shading: doNotUseGlobalStyle
          ? textStyle.shading
          : textStyle.shading ?? globalTextStyle.shading,
      smallCaps: doNotUseGlobalStyle
          ? textStyle.smallCaps
          : textStyle.smallCaps ?? globalTextStyle.smallCaps,
      strike: doNotUseGlobalStyle
          ? textStyle.strike
          : textStyle.strike ?? globalTextStyle.strike,
      textArt: doNotUseGlobalStyle
          ? textStyle.textArt
          : textStyle.textArt ?? globalTextStyle.textArt,
      underline: doNotUseGlobalStyle
          ? textStyle.underline
          : textStyle.underline ?? globalTextStyle.underline,
      underlineColor: doNotUseGlobalStyle
          ? textStyle.underlineColor
          : textStyle.underlineColor ?? globalTextStyle.underlineColor,
      vanish: doNotUseGlobalStyle
          ? textStyle.vanish
          : textStyle.vanish ?? globalTextStyle.vanish,
      vertAlign: doNotUseGlobalStyle
          ? textStyle.vertAlign
          : textStyle.vertAlign ?? globalTextStyle.vertAlign,
      spacing: doNotUseGlobalStyle
          ? textStyle.textSpacing
          : textStyle.textSpacing ?? globalTextStyle.textSpacing,
      stretchOrCompressInPercentage: doNotUseGlobalStyle
          ? textStyle.stretchOrCompressInPercentage
          : textStyle.stretchOrCompressInPercentage ??
              globalTextStyle.stretchOrCompressInPercentage,
      kern: doNotUseGlobalStyle
          ? textStyle.kern
          : textStyle.kern ?? globalTextStyle.kern,
      fitText: doNotUseGlobalStyle
          ? textStyle.fitText
          : textStyle.fitText ?? globalTextStyle.fitText,
      effect: doNotUseGlobalStyle
          ? textStyle.textEffect
          : textStyle.textEffect ?? globalTextStyle.textEffect,
    );
  }

  List<String> _getOpenAndCloseHyperlinkTags(String hyperlinkTo) {
    String openHyperlink = '';
    String closeHyperlink = '';
    if (!hyperlinkTo.startsWith('#')) {
      _packager.addHyperlink(hyperlinkTo);
    }
    closeHyperlink = '</w:hyperlink>';
    final String linkHyperlink = hyperlinkTo.startsWith('#')
        ? 'w:anchor="${hyperlinkTo.substring(1)}"'
        : 'r:id="rId${_packager.rIdCount - 1}"';
    openHyperlink = '<w:hyperlink $linkHyperlink>';
    // openHyperlink = '<w:hyperlink r:id="rId${_packager.rIdCount - 1}">';
    return [openHyperlink, closeHyperlink];
  }

  List<String> _getOpenAndCloseBookmarkTags(String name) {
    final String openTag =
        '<w:bookmarkStart w:id="$_anchorCounter" w:name="$name"/>';
    final String closeTag = '<w:bookmarkEnd w:id="$_anchorCounter"/>';
    bookmarks.add(name);
    _anchorCounter++;
    return [openTag, closeTag];
  }

  /// AddText adds lines of text to the document.
  /// By default, it uses the global text styling. This can be changed by providing a [textStyle] and/or setting [doNotUseGlobalStyle] to false.
  /// If the text is supposed to be a list item, pass a textStyle which has a numberingList (which decides the general look of the list) and optionally a numberLevelInList (which decides the depth and indentation of the item; 0 by default).
  ///
  /// This function always adds a new paragraph to the document.
  ///
  /// If [lineOrPageBreak] is given, then a LineBreak will be added at the end of the text.
  /// If the textStyle has a non-empty Tabs list, then a tab can be added in front of the text by setting [addTab] to true.
  ///
  /// In addText, the fields complexField, hyperlinkTo and setBookmarkName are mutually exclusive.
  ///
  /// If given, [complexField] adds a complex field, such as page and date, e.g. ComplexField(instructions: 'PAGE') instructs the word processor to insert the current page number. Or for an internal hyperlink: ComplexField(instructions: ' HYPERLINK \\l &quot;bookmark1&quot; ').
  /// If [hyperlinkTo] is not null or empty (e.g. "https://www.flutter.dev"), then the text becomes a hyperlink. If the hyperlinkTo has an internal destination (a bookmark) then prefix the destination with a '#' sign, e.g. "#bookmark1". (The globalTextStyle's hyperlinkTo is always ignored)
  /// If provided, [setBookmarkName] creates a bookmark with the given name which can be used as a target for internal hyperlinks.
  void addText(
    String text, {
    TextStyle textStyle,
    bool doNotUseGlobalStyle = false,
    LineBreak lineOrPageBreak,
    bool addTab = false,
    String hyperlinkTo,
    String setBookmarkName,
    ComplexField complexField,
  }) {
    if (!_bufferClosed) {
      _docx.write(_getCachedAddText(
        text,
        textStyle: textStyle,
        doNotUseGlobalStyle: doNotUseGlobalStyle,
        lineOrPageBreak: lineOrPageBreak,
        addTab: addTab,
        hyperlinkTo: hyperlinkTo,
        complexField: complexField,
        setBookmarkName: setBookmarkName,
        // textFrame: textFrame,
        // numberingList: numberingList,
        // numberLevelInList: numberLevelInList,
      ));
      // _addToCharCounters(text);
      // _parCount++;
    }
  }

  /// Writes XML text to cache; This data can then either be written to a stringbuffer or used as content for table cells.
  String _getCachedAddText(
    String text, {
    TextStyle textStyle,
    LineBreak lineOrPageBreak,
    bool addTab = false,
    String hyperlinkTo,
    String setBookmarkName,
    ComplexField complexField,
    bool doNotUseGlobalStyle = false,
  }) {
    final StringBuffer cached = StringBuffer();
    final String tab =
        globalTextStyle.tabs != null && addTab ? '<w:r><w:tab/></w:r>' : '';
    final String lineBreak =
        lineOrPageBreak != null ? lineOrPageBreak.getXml() : '';
    final TextStyle style =
        textStyle ?? (doNotUseGlobalStyle ? TextStyle() : globalTextStyle);
    final String ppr = _getParagraphStyleAsString(
        textStyle: style, doNotUseGlobalStyle: doNotUseGlobalStyle);

    if (complexField == null ||
        complexField.instructions == null ||
        complexField.includeSeparate == null) {
      bool styleAsHyperlink = false;
      String closeHyperlink = '';
      String openHyperlink = '';
      if (hyperlinkTo != null && hyperlinkTo.isNotEmpty) {
        final tags = _getOpenAndCloseHyperlinkTags(hyperlinkTo);
        openHyperlink = tags[0];
        closeHyperlink = tags[1];
        styleAsHyperlink = true;
      }
      if (openHyperlink.isEmpty &&
          setBookmarkName != null &&
          setBookmarkName.isNotEmpty) {
        final tags = _getOpenAndCloseBookmarkTags(setBookmarkName);
        openHyperlink = tags[0];
        closeHyperlink = tags[1];
      }

      cached.writeAll(<String>[
        '<w:p>$ppr',
        tab,
        openHyperlink,
        '<w:r>${_getTextStyleAsString(style: style, styleAsHyperlink: styleAsHyperlink, doNotUseGlobalStyle: doNotUseGlobalStyle)}${text.startsWith(' ') || text.endsWith(' ') ? '<w:t xml:space="preserve">' : '<w:t>'}$text</w:t></w:r>$lineBreak$closeHyperlink</w:p>'
      ]);
    } else {
      cached.writeAll(<String>[
        '<w:p>$ppr',
        tab,
        '<w:r>${_getTextStyleAsString(style: style, doNotUseGlobalStyle: doNotUseGlobalStyle)}<w:t xml:space="preserve">$text</w:t></w:r><w:r>${complexField.getXml()}</w:r>$lineBreak</w:p>'
      ]);
    }
    _debugString = cached.toString();
    return cached.toString();
  }

  /// AddMixedText adds lines of text that do NOT have the same styling as each other or with the global text style.
  /// Unless [doNotUseTextGlobalStyle] is set to true, global text styling will be used for any values not provided (i.e. null values) by custom styling rules.
  /// Given lists should have equal lengths. Page style is optional and is used for custom section styling. If [doNotUsePageGlobalStyle] is set to false, global page styling will be used for any values not provided (i.e. null values) by the custom page styling.
  /// Note that the last paragraph of the document should NOT have any custom section styling!
  /// Except for HighlightColor, colors are in 'RRGGBB' format.
  ///
  /// Lists can hold null values. For text: null implies no text; for textStyles: null implies use of globalTextStyle. [textAlignment] for the entire text should be set as an argument in the function instead of in the individual textStyles.
  ///
  /// This function always adds a new paragraph to the document.
  ///
  /// If [lineOrPageBreak] is given, then a LineBreak will be added after every item (if [addBreakAfterEveryItem] is true) or only after the last item on the [text] list (if [addBreakAfterEveryItem] is false, which is default).
  /// If globalTextStyle has a non-empty Tabs list, then a tab can be added in front of the first text item by setting [addTab] to true.
  ///
  /// If the custom textstyles contain a hyperlinkTo value that is not null or empty, then the text will be a hyperlink and direct to [hyperlinkTo]. The global textstyle's hyperlinkTo is always ignored.
  ///
  /// If given, [complexFields] should have the same length as [text] and includes complex fields, such as page and date. For example, ComplexField(instructions: 'PAGE') instructs the word processor to insert the current page number. A complexField cannot be added if hyperlinkTo is not null.
  ///
  /// Mixed text can be displayed within a [textFrame] or as a list item in [numberingList] with depth [numberLevelInList].
  void addMixedText(
    List<String> text,
    List<TextStyle> textStyles, {
    TextAlignment textAlignment,
    PageStyle pageStyle,
    bool doNotUseGlobalTextStyle = false,
    bool doNotUseGlobalPageStyle = true,
    LineBreak lineOrPageBreak,
    bool addBreakAfterEveryItem = false,
    bool addTab = false,
    List<ComplexField> complexFields,
    TextFrame textFrame,
    String setBookmarkName,
    NumberingList numberingList,
    int numberLevelInList,
  }) {
    if (!_bufferClosed && text.isNotEmpty && text.length == textStyles.length) {
      _docx.write(_getCachedAddMixedText(
        text,
        textStyles,
        textAlignment: textAlignment,
        pageStyle: pageStyle,
        doNotUseGlobalTextStyle: doNotUseGlobalTextStyle,
        doNotUseGlobalPageStyle: doNotUseGlobalPageStyle,
        lineOrPageBreak: lineOrPageBreak,
        addBreakAfterEveryItem: addBreakAfterEveryItem,
        addTab: addTab,
        complexFields: complexFields,
        textFrame: textFrame,
        setBookmarkName: setBookmarkName,
        numberingList: numberingList,
        numberLevelInList: numberLevelInList,
      ));
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
    TextAlignment textAlignment,
    TextFrame textFrame,
    String setBookmarkName,
    NumberingList numberingList,
    int numberLevelInList,
  }) {
    final StringBuffer d = StringBuffer();
    d.write('<w:p><w:pPr>');
    if (pageStyle != null) {
      d.write(_getPageStyleAsString(
          style: pageStyle, doNotUseGlobalStyle: doNotUseGlobalPageStyle));
    }
    d.write(_getParagraphStyleAsString(
            textStyle: TextStyle(
              textAlignment: textAlignment,
              textFrame: textFrame,
              numberingList: numberingList,
              numberLevelInList: numberLevelInList,
            ),
            doNotUseGlobalStyle: doNotUseGlobalTextStyle)
        .replaceFirst('<w:pPr>', ''));

    final String multiBreak = lineOrPageBreak != null && addBreakAfterEveryItem
        ? lineOrPageBreak.getXml()
        : '';

    String openBookmarkTag = '';
    String closeBookmarkTag = '';
    if (setBookmarkName != null && setBookmarkName.isNotEmpty) {
      final tags = _getOpenAndCloseBookmarkTags(setBookmarkName);
      openBookmarkTag = tags[0];
      closeBookmarkTag = tags[1];
    }

    if (globalTextStyle.tabs != null && addTab) {
      d.write('<w:r><w:tab/></w:r>');
    }
    d.write(openBookmarkTag);
    final List<ComplexField> cf =
        complexFields != null && complexFields.length == text.length
            ? complexFields
            : List<ComplexField>.generate(text.length, (index) => null);

    for (int i = 0; i < text.length; i++) {
      final String t = text[i] ?? '';
      final ComplexField f = cf[i];
      if (f == null || f.instructions == null || f.includeSeparate == null) {
        bool styleAsHyperlink = false;
        String closeHyperlink = '';
        String openHyperlink = '';
        if (textStyles[i] != null &&
            textStyles[i].hyperlinkTo != null &&
            textStyles[i].hyperlinkTo.isNotEmpty) {
          final tags = _getOpenAndCloseHyperlinkTags(textStyles[i].hyperlinkTo);
          openHyperlink = tags[0];
          closeHyperlink = tags[1];
          styleAsHyperlink = true;
        }
        d.writeAll(<String>[
          openHyperlink,
          '<w:r>',
          _getTextStyleAsString(
              styleAsHyperlink: styleAsHyperlink,
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

    d.write('$closeBookmarkTag$singleBreak</w:p>');

    _debugString = d.toString();
    return d.toString();
  }

  /// When the table has been finalized, attach it to the document at the cursor's current position. A table should contain a list of column widths and have at least 1 table row. Rows can contain zero or more table cells.
  ///
  /// Similar to headers and footers, tables will only become visible after being attached to the document.
  void attachTable(Table table) {
    _debugString = table.getXml();
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

  /// A complete table can be inserted into a table cell.
  void insertTableInTableCell(
          {@required Table table, @required TableCell tableCell}) =>
      tableCell.xmlContent = table.getXml();

  /// Inserts text into the given [tableCell]. The result is similar as if addText() was used.
  ///
  /// See addText() for more info.
  void insertTextInTableCell({
    @required TableCell tableCell,
    @required String text,
    TextStyle textStyle,
    LineBreak lineOrPageBreak,
    bool addTab = false,
    String hyperlinkTo,
    String setBookmarkName,
    ComplexField complexField,
    bool doNotUseGlobalStyle = false,
  }) {
    if (!_bufferClosed) {
      tableCell.xmlContent = _getCachedAddText(
        text,
        textStyle: textStyle,
        lineOrPageBreak: lineOrPageBreak,
        addTab: addTab,
        hyperlinkTo: hyperlinkTo,
        setBookmarkName: setBookmarkName,
        complexField: complexField,
        doNotUseGlobalStyle: doNotUseGlobalStyle,
      );
      // _addToCharCounters(text);
      // _parCount++;
    }
  }

  /// Inserts mixed text into the given [tableCell]. The result is similar as if addMixedText() was used.
  ///
  /// See addMixedText() for more info.
  void insertMixedTextInTableCell({
    @required TableCell tableCell,
    @required List<String> text,
    TextAlignment textAlignment,
    List<TextStyle> textStyles,
    PageStyle pageStyle,
    bool doNotUseGlobalTextStyle = false,
    bool doNotUseGlobalPageStyle = true,
    LineBreak lineOrPageBreak,
    bool addBreakAfterEveryItem = false,
    bool addTab = false,
    List<ComplexField> complexFields,
    TextFrame textFrame,
    String setBookmarkName,
    NumberingList numberingList,
    int numberLevelInList,
  }) {
    if (!_bufferClosed && text.isNotEmpty && text.length == textStyles.length) {
      tableCell.xmlContent = _getCachedAddMixedText(
        text,
        textStyles,
        textAlignment: textAlignment,
        pageStyle: pageStyle,
        doNotUseGlobalTextStyle: doNotUseGlobalTextStyle,
        doNotUseGlobalPageStyle: doNotUseGlobalPageStyle,
        lineOrPageBreak: lineOrPageBreak,
        addBreakAfterEveryItem: addBreakAfterEveryItem,
        addTab: addTab,
        complexFields: complexFields,
        setBookmarkName: setBookmarkName,
        textFrame: textFrame,
        numberingList: numberingList,
        numberLevelInList: numberLevelInList,
      );
      // _addToCharCounters(t);
      // _parCount++;
    }
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
    AnchorImageHorizontalAlignment anchorImageHorizontalAlignment =
        AnchorImageHorizontalAlignment.center,
    AnchorImageVerticalAlignment anchorImageVerticalAlignment =
        AnchorImageVerticalAlignment.top,
    String alternativeTextForImage = '',
    bool noChangeAspect = false,
    bool noChangeArrowheads = false,
    bool noMove = false,
    bool noResize = false,
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

    if (widthEMU < 1 || widthEMU > max || heightEMU < 1 || heightEMU > max) {
      return;
    }

    if (!_bufferClosed) {
      _docx.write(_getCachedAnchorImage(
        imageFile,
        widthEMU,
        heightEMU,
        horizontalOffsetEMU,
        verticalOffsetEMU,
        distB: distB,
        distL: distL,
        distR: distR,
        distT: distT,
        addBreakAfterEveryItem: addBreakAfterEveryItem,
        addTab: addTab,
        allowOverlap: allowOverlap,
        alternativeTextForImage: alternativeTextForImage,
        anchorImageAreaWrap: anchorImageAreaWrap,
        anchorImageHorizontalAlignment: anchorImageHorizontalAlignment,
        anchorImageVerticalAlignment: anchorImageVerticalAlignment,
        anchorImageTextWrap: anchorImageTextWrap,
        behindDocument: behindDocument,
        doNotUseGlobalPageStyle: doNotUseGlobalPageStyle,
        doNotUseGlobalTextStyle: doNotUseGlobalTextStyle,
        effectExtentB: effectExtentB,
        effectExtentL: effectExtentL,
        effectExtentR: effectExtentR,
        effectExtentT: effectExtentT,
        encloseInParagraph: true,
        flipImageHorizontal: flipImageHorizontal,
        flipImageVertical: flipImageVertical,
        horizontalPositionRelativeBase: horizontalPositionRelativeBase,
        hyperlinkTo: hyperlinkTo,
        imageTextStyle: imageTextStyle,
        layoutInCell: layoutInCell,
        lineOrPageBreak: lineOrPageBreak,
        locked: locked,
        noChangeArrowheads: noChangeArrowheads,
        noChangeAspect: noChangeAspect,
        noMove: noMove,
        noResize: noResize,
        noSelect: noSelect,
        pageStyle: pageStyle,
        relativeHeight: relativeHeight,
        rotateInEMU: rotateInEMU,
        simplePos: simplePos,
        simplePosX: simplePosX,
        simplePosY: simplePosY,
        text: text,
        textStyles: textStyles,
        verticalPositionRelativeBase: verticalPositionRelativeBase,
      ));
      if (text != null) {
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

  String _getCachedAnchorImage(
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
    AnchorImageHorizontalAlignment anchorImageHorizontalAlignment =
        AnchorImageHorizontalAlignment.center,
    AnchorImageVerticalAlignment anchorImageVerticalAlignment =
        AnchorImageVerticalAlignment.top,
    String alternativeTextForImage = '',
    bool noChangeAspect = false,
    bool noChangeArrowheads = false,
    bool noMove = false,
    bool noResize = false,
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
    bool encloseInParagraph = true,
  }) {
    final style =
        imageTextStyle ?? TextStyle(textAlignment: TextAlignment.center);

    final String path = imageFile.path;
    final String suffix = path.substring(path.lastIndexOf('.') + 1);
    String result;
    if (mimeTypes.containsKey(suffix)) {
      final bool saved = _packager.addImageFile(imageFile, suffix);
      if (saved) {
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
            ? '${_getParagraphStyleAsString(textStyle: style, doNotUseGlobalStyle: doNotUseGlobalTextStyle)}'
            : '';
        const String openR = '<w:r>';
        const String closeR = '</w:r>';

        result =
            '$openParagraph$openPpr$openR<w:drawing><wp:anchor behindDoc="$behindDocument" distT="$distT" distB="$distB" distL="$distL" distR="$distR" simplePos="$simplePos" locked="$locked" layoutInCell="$layoutInCell" allowOverlap="$allowOverlap" relativeHeight="$relativeHeight"><wp:simplePos x="$simplePosX" y="$simplePosY" /><wp:positionH relativeFrom="${getValueFromEnum(horizontalPositionRelativeBase)}"><wp:align>${getValueFromEnum(anchorImageHorizontalAlignment)}</wp:align><wp:posOffset>$horizontalOffsetEMU</wp:posOffset></wp:positionH><wp:positionV relativeFrom="${getValueFromEnum(verticalPositionRelativeBase)}"><wp:align>${getValueFromEnum(anchorImageVerticalAlignment)}</wp:align><wp:posOffset>$verticalOffsetEMU</wp:posOffset></wp:positionV><wp:extent cx="$widthEMU" cy="$heightEMU"/><wp:effectExtent l="$effectExtentL" t="$effectExtentT" r="$effectExtentR" b="$effectExtentB"/><wp:${getValueFromEnum(anchorImageAreaWrap)} wrapText="${getValueFromEnum(anchorImageTextWrap)}"/><wp:docPr id="$mediaIdCount" name="Image$mediaIdCount" descr="$alternativeTextForImage">$hyperlink</wp:docPr><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="$noChangeAspect" noMove="$noMove" noResize="$noResize" noSelect="$noSelect"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="$mediaIdCount" name="Image$mediaIdCount" descr="$alternativeTextForImage"></pic:cNvPr><pic:cNvPicPr><a:picLocks noChangeAspect="$noChangeAspect" noMove="$noMove" noResize="$noResize" noSelect="$noSelect" noChangeArrowheads="$noChangeArrowheads"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId$mediaIdCount"></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr bwMode="auto">$xfrm</pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>$closeR$closeParagraph';
      } else {
        result = '';
      }
    } else {
      result = '';
    }
    _debugString = result;
    return result;
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
    bool noChangeArrowheads = false,
    bool noChangeAspect = false,
    bool noMove = false,
    bool noResize = false,
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
        noChangeArrowheads: noChangeArrowheads,
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
    bool noChangeArrowheads = false,
    bool noChangeAspect = false,
    bool noMove = false,
    bool noResize = false,
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
        noChangeArrowheads: noChangeArrowheads,
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

  /// Creates an inline image with a caption inside a 1x1 table. The [tableProperties] can be set to customize the appearance of the table, such as changing the border colors and widths. LibreOffice/OpenOffice converts an inline image (called "anchor as character") in a table to an anchor image (called "anchor to character/paragraph") for historic reasons, but this should have little effect on the end result.
  ///
  /// Other ways to combine image with text: use addAnchorImage or use addImage followed or preceeded with addText/addMixedText.
  ///
  /// [tableWidthInTwips] determines the width of the table (measured in twips). If omitted, the width of the image will be used as table width.
  /// [captionAppearsBelowImage] determines if the caption is displayed above (false) or below (true; default) the image.
  /// [hyperlinkTo] applies to the image, not the caption. If the caption needs to be a hyperlink, you can use the textStyle.hyperlinkTo.
  void addImageWithCaption(
    File imageFile,
    int widthEMU,
    int heightEMU, {
    int tableWidthInTwips,
    bool captionAppearsBelowImage = true,
    List<String> text,
    List<TextStyle> textStyles,
    String alternativeTextForImage = '',
    bool noChangeArrowheads = false,
    bool noChangeAspect = false,
    bool noMove = false,
    bool noResize = false,
    bool noSelect = false,
    String hyperlinkTo,
    int effectExtentL = 0,
    int effectExtentT = 0,
    int effectExtentR = 0,
    int effectExtentB = 0,
    bool flipImageHorizontal = false,
    bool flipImageVertical = false,
    int rotateInEMU = 0,
    TableProperties tableProperties,
    TextAlignment textAlignment = TextAlignment.center,
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
        noChangeArrowheads: noChangeArrowheads,
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
        textAlignment: textAlignment,
        complexFields: complexFields,
        doNotUseGlobalPageStyle: doNotUseGlobalPageStyle,
        doNotUseGlobalTextStyle: doNotUseGlobalTextStyle,
        addTab: addTab,
      );

      final TableRow row = TableRow(
        tableRowProperties: TableRowProperties(
            tableCellSpacingWidthType: PreferredWidthType.auto),
        tableCells: [
          TableCell(
            tableCellProperties: TableCellProperties(
              restartVMerge: true,
              verticalAlignment: tableCellVerticalAlignment,
            ),
            xmlContent: captionAppearsBelowImage
                ? '$xmlImage$xmlCaption'
                : '$xmlCaption$xmlImage',
          )
        ],
      );

      final Table t = Table(
        gridColumnWidths: [tableWidthInTwips ?? convertEMUToTwips(widthEMU)],
        tableProperties: tableProperties ??
            TableProperties(
              tableTextAlignment: TableTextAlignment.center,
              tableCellSpacing: '0',
              tableCellSpacingType: PreferredWidthType.nil,
              tableCellMargins: [
                TableCellMargin(
                  isNil: true,
                  tableCellSide: TableCellSide.right,
                ),
                TableCellMargin(
                  isNil: true,
                  tableCellSide: TableCellSide.left,
                ),
                TableCellMargin(
                  isNil: true,
                  tableCellSide: TableCellSide.top,
                ),
                TableCellMargin(
                  isNil: true,
                  tableCellSide: TableCellSide.bottom,
                ),
              ],
            ),
        tableRows: [row],
      );

      attachTable(t);
    }
  }

  /// Inserts inline image into the table cell.
  void insertImageInTableCell({
    @required TableCell tableCell,
    @required File imageFile,
    @required int widthEMU,
    @required int heightEMU,
    String alternativeTextForImage = '',
    bool noChangeArrowheads = false,
    bool noChangeAspect = false,
    bool noMove = false,
    bool noResize = false,
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
        noChangeArrowheads: noChangeArrowheads,
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

  /// Inserts multiple inline images into the table cell.
  void insertImagesInTableCell({
    @required TableCell tableCell,
    @required List<File> imageFiles,
    @required int widthEMU,
    @required int heightEMU,
    List<String> alternativeTextListForImage,
    List<String> hyperlinksTo,
    bool noChangeArrowheads = false,
    bool noChangeAspect = false,
    bool noMove = false,
    bool noResize = false,
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
        noChangeArrowheads: noChangeArrowheads,
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
    _debugString = s.toString();
    tableCell.xmlContent = s.toString();
  }

  void _insertInlineImage(
    File imageFile,
    int widthEMU,
    int heightEMU, {
    String alternativeTextForImage = '',
    bool noChangeArrowheads = false,
    bool noChangeAspect = false,
    bool noMove = false,
    bool noResize = false,
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
        noChangeArrowheads: noChangeArrowheads,
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
      bool noChangeArrowheads = false,
      bool noChangeAspect = false,
      bool noMove = false,
      bool noResize = false,
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
            '$openParagraph$openPpr$openR<w:drawing><wp:inline><wp:extent cx="$widthEMU" cy="$heightEMU"/><wp:effectExtent l="$effectExtentL" t="$effectExtentT" r="$effectExtentR" b="$effectExtentB"/><wp:docPr id="$mediaIdCount" name="Image$mediaIdCount" descr="$alternativeTextForImage">$hyperlink</wp:docPr><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="$noChangeAspect" noMove="$noMove" noResize="$noResize" noSelect="$noSelect"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="$mediaIdCount" name="Image$mediaIdCount" descr="$alternativeTextForImage"></pic:cNvPr><pic:cNvPicPr><a:picLocks noChangeAspect="$noChangeAspect" noMove="$noMove" noResize="$noResize" noSelect="$noSelect" noChangeArrowheads="$noChangeArrowheads"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId$mediaIdCount"></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr bwMode="auto">$xfrm</pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>$closeR$closeParagraph');
      }
    }
    _debugString = d.toString();
    return d.toString();
  }

  /// Attaches a comment at the cursor's current position in the document with the provided [message]; the fields [author] and [initials] can also be given.
  ///
  /// If provided, [timestamp] should be in the UTC format, e.g. "2020-07-28T20:32:56Z". If [timestamp] is omitted, the current date and time is used.
  void attachComment({
    String author = 'Unknown',
    String message = '',
    String timestamp,
    String initials = '',
  }) {
    final String datetime =
        timestamp ?? getUtcTimestampFromDateTime(DateTime.now());
    comments[_commentCounter] =
        '<w:comment w:id="$_commentCounter" w:author="$author" w:date="$datetime" w:initials="$initials"><w:p><w:pPr><w:rPr></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint="default"/></w:rPr><w:t>$message</w:t></w:r></w:p></w:comment>';
    final comment =
        '<w:p><w:r><w:commentReference w:id="$_commentCounter"/></w:r></w:p>';
    _debugString = comment;
    _docx.write(comment);
    _commentCounter++;
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
        if (_includeNumberingXml) {
          _packager.addNumberingList();
        }
        if (addEmptyParagraphAtEndOfDocument) {
          addText('', textStyle: TextStyle(), doNotUseGlobalStyle: true);
        }
        _docx.write(_getPageStyleAsString(style: globalPageStyle));
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
        customNumberingXml: _customNumberingXml,
        includeNumberingXml: _includeNumberingXml,
        comments: comments,
        hyperlinkStylingXml: hyperlinkTextStyle != null
            ? _getTextStyleAsString(
                style: hyperlinkTextStyle, doNotUseGlobalStyle: true)
            : null,
      );
      return f;
    } catch (e) {
      throw StateError('DOCX BUILDER ERROR: $e');
    }
  }

  /// Deletes all data and cache, so a new document can be created (or when DocXBuilder is no longer needed).
  /// This function also destroys the source file produced by createDocxFile so be sure to save or process the source file before calling clear.
  void clear({bool resetBuffer = true}) {
    _docx.clear();
    _debugString = null;
    _documentBackgroundColor = null;
    bookmarks.clear();
    comments.clear();
    hyperlinkTextStyle = null;
    globalTextStyle = TextStyle();
    globalPageStyle = PageStyle().getDefaultPageStyle();
    _firstPageHeader = null;
    _oddPageHeader = null;
    _evenPageHeader = null;
    _firstPageFooter = null;
    _oddPageFooter = null;
    _evenPageFooter = null;
    _insertHeadersAndFootersInThisSection = false;
    _customNumberingXml = null;
    _includeNumberingXml = false;
    // _charCount = 0;
    // _charCountWithSpaces = 0;
    // _parCount = 0;
    _headerCounter = 1;
    _footerCounter = 1;
    _anchorCounter = 0;
    _commentCounter = 0;
    _packager.destroyCache();
    if (resetBuffer) {
      _initDocX();
    }
  }

  String debugBuffer() => _debugString;
}
