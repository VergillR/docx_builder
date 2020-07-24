import 'dart:io';

import 'package:docx_builder/docx_builder.dart';
import 'package:docx_builder/src/utils/constants/mimetypes.dart';
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

  DocxTextStyle _globalDocxTextStyle = DocxTextStyle();
  DocxTextStyle get globalDocxTextStyle => _globalDocxTextStyle;
  DocxPageStyle _globalDocxPageStyle = DocxPageStyle().getDefaultPageStyle();
  DocxPageStyle get globalDocxPageStyle => _globalDocxPageStyle;

  static String mimetype =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

  final int emuWidthA4Pct10 = 756000;
  final int emuWidthA4Pct20 = 1512000;
  final int emuWidthA4Pct30 = 2268000;
  final int emuWidthA4Pct40 = 3024000;
  final int emuWidthA4Pct50 = 3780000;
  final int emuWidthA4Pct60 = 4536000;
  final int emuWidthA4Pct70 = 5292000;
  final int emuWidthA4Pct80 = 6048000;
  final int emuWidthA4Pct90 = 6804000;
  final int emuWidthA4Pct100 = 7560000;

  final int emuHeightA4Pct10 = 1069200;
  final int emuHeightA4Pct20 = 2138400;
  final int emuHeightA4Pct30 = 3207600;
  final int emuHeightA4Pct40 = 4276800;
  final int emuHeightA4Pct50 = 5346000;
  final int emuHeightA4Pct60 = 6415200;
  final int emuHeightA4Pct70 = 7484400;
  final int emuHeightA4Pct80 = 8553600;
  final int emuHeightA4Pct90 = 9622800;
  final int emuHeightA4Pct100 = 10692000;

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
  String _getDocxPageStyleAsString(
      {DocxPageStyle style, bool doNotUseGlobalStyle = false}) {
    final DocxPageStyle pageStyle = style ?? DocxPageStyle();
    return _b.SectPr.getSpr(
      cols: doNotUseGlobalStyle
          ? pageStyle.cols
          : pageStyle.cols ?? _globalDocxPageStyle.cols,
      colsHaveSeparator: doNotUseGlobalStyle
          ? pageStyle.colsHaveSeparator
          : pageStyle.colsHaveSeparator ??
              _globalDocxPageStyle.colsHaveSeparator,
      colsHaveEqualWidth: doNotUseGlobalStyle
          ? pageStyle.colsHaveEqualWidth
          : pageStyle.colsHaveEqualWidth ??
              _globalDocxPageStyle.colsHaveEqualWidth,
      sectType: doNotUseGlobalStyle
          ? pageStyle.sectType
          : pageStyle.sectType ?? _globalDocxPageStyle.sectType,
      pageSzHeight: doNotUseGlobalStyle
          ? pageStyle.pageSzHeight
          : pageStyle.pageSzHeight ?? _globalDocxPageStyle.pageSzHeight,
      pageSzWidth: doNotUseGlobalStyle
          ? pageStyle.pageSzWidth
          : pageStyle.pageSzWidth ?? _globalDocxPageStyle.pageSzWidth,
      pageOrientation: doNotUseGlobalStyle
          ? pageStyle.pageOrientation
          : pageStyle.pageOrientation ?? _globalDocxPageStyle.pageOrientation,
      sectVerticalAlign: doNotUseGlobalStyle
          ? pageStyle.sectVerticalAlign
          : pageStyle.sectVerticalAlign ??
              _globalDocxPageStyle.sectVerticalAlign,
      pageBorders: doNotUseGlobalStyle
          ? pageStyle.pageBorders
          : pageStyle.pageBorders ?? _globalDocxPageStyle.pageBorders,
      pageBorderDisplay: doNotUseGlobalStyle
          ? pageStyle.pageBorderDisplay
          : pageStyle.pageBorderDisplay ??
              _globalDocxPageStyle.pageBorderDisplay,
      pageBorderOffsetBasedOnText: doNotUseGlobalStyle
          ? pageStyle.pageBorderOffsetBasedOnText
          : pageStyle.pageBorderOffsetBasedOnText ??
              _globalDocxPageStyle.pageBorderOffsetBasedOnText,
      pageBorderIsRenderedAboveText: doNotUseGlobalStyle
          ? pageStyle.pageBorderIsRenderedAboveText
          : pageStyle.pageBorderIsRenderedAboveText ??
              _globalDocxPageStyle.pageBorderIsRenderedAboveText,
      pageNumberingFormat: doNotUseGlobalStyle
          ? pageStyle.pageNumberingFormat
          : pageStyle.pageNumberingFormat ??
              _globalDocxPageStyle.pageNumberingFormat,
      pageNumberingStart: doNotUseGlobalStyle
          ? pageStyle.pageNumberingStart
          : pageStyle.pageNumberingStart ??
              _globalDocxPageStyle.pageNumberingStart,
      lineNumbering: doNotUseGlobalStyle
          ? pageStyle.lineNumbering
          : pageStyle.lineNumbering ?? _globalDocxPageStyle.lineNumbering,
      pageMargin: doNotUseGlobalStyle
          ? pageStyle.pageMargin
          : pageStyle.pageMargin ?? _globalDocxPageStyle.pageMargin,
    );
  }

  // ignore: use_setters_to_change_properties
  /// Sets the global style for text. Except for HighlightColor, colors are always 'RRGGBB' format.
  /// The global text style can be overridden by addMixedText with its own styling rules.
  /// The hyperlinkTo property of the global text style is not used by addText or addMixedText.
  void setGlobalDocxTextStyle(DocxTextStyle textStyle) =>
      _globalDocxTextStyle = textStyle;

  /// Obtain the XML string of the paragraph style, such as text alignment.
  /// If no style is given, then the globalDocxTextStyle is used (unless [doNotUseGlobalStyle] is set to true).
  String _getParagraphStyleAsString(
      {DocxTextStyle textStyle, bool doNotUseGlobalStyle = false}) {
    final DocxTextStyle style = textStyle ?? DocxTextStyle();
    return _b.Ppr.getPpr(
      keepLines: doNotUseGlobalStyle
          ? style.keepLines
          : style.keepLines ?? _globalDocxTextStyle.keepLines,
      keepNext: doNotUseGlobalStyle
          ? style.keepNext
          : style.keepNext ?? _globalDocxTextStyle.keepNext,
      paragraphBorderOnAllSides: doNotUseGlobalStyle
          ? style.paragraphBorderOnAllSides
          : style.paragraphBorderOnAllSides ??
              _globalDocxTextStyle.paragraphBorderOnAllSides,
      paragraphBorders: doNotUseGlobalStyle
          ? style.paragraphBorders
          : style.paragraphBorders ?? _globalDocxTextStyle.paragraphBorders,
      paragraphIndent: doNotUseGlobalStyle
          ? style.paragraphIndent
          : style.paragraphIndent ?? _globalDocxTextStyle.paragraphIndent,
      paragraphShading: doNotUseGlobalStyle
          ? style.paragraphShading
          : style.paragraphShading ?? _globalDocxTextStyle.paragraphShading,
      spacing: doNotUseGlobalStyle
          ? style.paragraphSpacing
          : style.paragraphSpacing ?? _globalDocxTextStyle.paragraphSpacing,
      tabs: doNotUseGlobalStyle
          ? style.tabs
          : style.tabs ?? _globalDocxTextStyle.tabs,
      textAlignment: doNotUseGlobalStyle
          ? style.textAlignment
          : style.textAlignment ?? _globalDocxTextStyle.textAlignment,
      verticalTextAlignment: doNotUseGlobalStyle
          ? style.verticalTextAlignment
          : style.verticalTextAlignment ??
              _globalDocxTextStyle.verticalTextAlignment,
    );
  }

  /// Obtain the XML string of the text style.
  /// If no style is given, then the globalDocxTextStyle is used (unless [doNotUseGlobalStyle] is set to true).
  String _getTextStyleAsString(
      {DocxTextStyle style,
      bool doNotUseGlobalStyle = false,
      bool styleAsHyperlink = false}) {
    final DocxTextStyle textStyle = style ?? DocxTextStyle();
    return _b.Rpr.getRpr(
      styleAsHyperlink: styleAsHyperlink,
      bold: doNotUseGlobalStyle
          ? textStyle.bold
          : textStyle.bold ?? _globalDocxTextStyle.bold,
      caps: doNotUseGlobalStyle
          ? textStyle.caps
          : textStyle.caps ?? _globalDocxTextStyle.caps,
      doubleStrike: doNotUseGlobalStyle
          ? textStyle.doubleStrike
          : textStyle.doubleStrike ?? _globalDocxTextStyle.doubleStrike,
      fontColor: doNotUseGlobalStyle
          ? textStyle.fontColor
          : textStyle.fontColor ?? _globalDocxTextStyle.fontColor,
      fontName: doNotUseGlobalStyle
          ? textStyle.fontName
          : textStyle.fontName ?? _globalDocxTextStyle.fontName,
      fontNameComplexScript: doNotUseGlobalStyle
          ? textStyle.fontNameComplexScript
          : textStyle.fontNameComplexScript ??
              _globalDocxTextStyle.fontNameComplexScript,
      fontSize: doNotUseGlobalStyle
          ? textStyle.fontSize
          : textStyle.fontSize ?? _globalDocxTextStyle.fontSize,
      highlightColor: doNotUseGlobalStyle
          ? textStyle.highlightColor
          : textStyle.highlightColor ?? _globalDocxTextStyle.highlightColor,
      italic: doNotUseGlobalStyle
          ? textStyle.italic
          : textStyle.italic ?? _globalDocxTextStyle.italic,
      rtlText: doNotUseGlobalStyle
          ? textStyle.rtlText
          : textStyle.rtlText ?? _globalDocxTextStyle.rtlText,
      shading: doNotUseGlobalStyle
          ? textStyle.shading
          : textStyle.shading ?? _globalDocxTextStyle.shading,
      smallCaps: doNotUseGlobalStyle
          ? textStyle.smallCaps
          : textStyle.smallCaps ?? _globalDocxTextStyle.smallCaps,
      strike: doNotUseGlobalStyle
          ? textStyle.strike
          : textStyle.strike ?? _globalDocxTextStyle.strike,
      textArt: doNotUseGlobalStyle
          ? textStyle.textArt
          : textStyle.textArt ?? _globalDocxTextStyle.textArt,
      underline: doNotUseGlobalStyle
          ? textStyle.underline
          : textStyle.underline ?? _globalDocxTextStyle.underline,
      underlineColor: doNotUseGlobalStyle
          ? textStyle.underlineColor
          : textStyle.underlineColor ?? _globalDocxTextStyle.underlineColor,
      vanish: doNotUseGlobalStyle
          ? textStyle.vanish
          : textStyle.vanish ?? _globalDocxTextStyle.vanish,
      vertAlign: doNotUseGlobalStyle
          ? textStyle.vertAlign
          : textStyle.vertAlign ?? _globalDocxTextStyle.vertAlign,
      spacing: doNotUseGlobalStyle
          ? textStyle.textSpacing
          : textStyle.textSpacing ?? _globalDocxTextStyle.textSpacing,
      stretchOrCompressInPercentage: doNotUseGlobalStyle
          ? textStyle.stretchOrCompressInPercentage
          : textStyle.stretchOrCompressInPercentage ??
              _globalDocxTextStyle.stretchOrCompressInPercentage,
      kern: doNotUseGlobalStyle
          ? textStyle.kern
          : textStyle.kern ?? _globalDocxTextStyle.kern,
      fitText: doNotUseGlobalStyle
          ? textStyle.fitText
          : textStyle.fitText ?? _globalDocxTextStyle.fitText,
      effect: doNotUseGlobalStyle
          ? textStyle.textEffect
          : textStyle.textEffect ?? _globalDocxTextStyle.textEffect,
    );
  }

  String _getLineOrPageBreak(LineBreak b) =>
      '<w:r><w:br w:type="${b.lineBreakType != null ? getValueFromEnum(b.lineBreakType) : "textWrapping"}" w:clear="${b.lineBreakClearLocation != null ? getValueFromEnum(b.lineBreakClearLocation) : "none"}" /></w:r>';

  /// AddText adds lines of text to the document.
  /// It uses the global text styling defined by setGlobalDocxTextStyle.
  /// This function always adds a new paragraph to the document.
  ///
  /// If [lineOrPageBreak] is given, then a LineBreak will be added at the end of the text.
  /// If globalDocxTextStyle has a non-empty Tabs list, then a tab can be added in front of the text by setting [addTab] to true.
  ///
  /// If [hyperlinkTo] is not null or empty, then the text will be a hyperlink and direct to [hyperlinkTo]. The global textstyle's hyperlinkTo is always ignored.
  void addText(String text,
      {LineBreak lineOrPageBreak, bool addTab = false, String hyperlinkTo}) {
    final String tab =
        globalDocxTextStyle.tabs != null && addTab ? '<w:r><w:tab/></w:r>' : '';
    final String lineBreak =
        lineOrPageBreak != null ? _getLineOrPageBreak(lineOrPageBreak) : '';

    String closeHyperlink = '';
    String openHyperlink = '';
    if (hyperlinkTo != null && hyperlinkTo.isNotEmpty) {
      _packager.addHyperlink(hyperlinkTo);
      closeHyperlink = '</w:hyperlink>';
      openHyperlink = '<w:hyperlink r:id="rId${_packager.rIdCount - 1}">';
    }

    if (!_bufferClosed) {
      _docxstring.writeAll(<String>[
        '<w:p>${_getParagraphStyleAsString()}',
        tab,
        openHyperlink,
        '<w:r>${_getTextStyleAsString(styleAsHyperlink: openHyperlink.isNotEmpty)}${text.startsWith(' ') || text.endsWith(' ') ? '<w:t xml:space="preserve">' : '<w:t>'}$text</w:t></w:r>$lineBreak$closeHyperlink</w:p>'
      ]);
      _addToCharCounters(text);
      _parCount++;
    }
  }

  /// AddMixedText adds lines of text that do NOT have the same styling as each other or with the global text style.
  /// Unless [doNotUseTextGlobalStyle] is set to true, global text styling will be used for any values not provided (i.e. null values) by custom styling rules.
  /// Given lists should have equal lengths. Page style is optional and is used for custom section styling. If [doNotUsePageGlobalStyle] is set to false, global page styling will be used for any values not provided (i.e. null values) by the custom page styling.
  /// Note that the last paragraph of the document should NOT have any custom section styling!
  /// Except for HighlightColor, colors are always 'RRGGBB' format.
  ///
  /// Lists can hold null values. For text: null implies no text; for textStyles: null implies use of globalDocxTextStyle.
  ///
  /// This function always adds a new paragraph to the document.
  ///
  /// If [lineOrPageBreak] is given, then a LineBreak will be added after every item (if [addBreakAfterEveryItem] is true) or only after the last item on the [text] list (if [addBreakAfterEveryItem] is false, which is default).
  /// If globalDocxTextStyle has a non-empty Tabs list, then a tab can be added in front of the first text item by setting [addTab] to true.
  ///
  /// If the custom textstyles contain a hyperlinkTo value that is not null or empty, then the text will be a hyperlink and direct to [hyperlinkTo]. The global textstyle's hyperlinkTo is always ignored.
  void addMixedText(
    List<String> text,
    List<DocxTextStyle> textStyles, {
    DocxPageStyle pageStyle,
    bool doNotUseGlobalTextStyle = false,
    bool doNotUseGlobalPageStyle = true,
    LineBreak lineOrPageBreak,
    bool addBreakAfterEveryItem = false,
    bool addTab = false,
  }) {
    if (!_bufferClosed && text.isNotEmpty && text.length == textStyles.length) {
      _docxstring.write('<w:p><w:pPr>');
      if (pageStyle != null) {
        _docxstring.write(_getDocxPageStyleAsString(
            style: pageStyle, doNotUseGlobalStyle: doNotUseGlobalPageStyle));
      }
      _docxstring
          .write(_getParagraphStyleAsString().replaceFirst('<w:pPr>', ''));

      final String multiBreak =
          lineOrPageBreak != null && addBreakAfterEveryItem
              ? _getLineOrPageBreak(lineOrPageBreak)
              : '';

      if (globalDocxTextStyle.tabs != null && addTab) {
        _docxstring.write('<w:r><w:tab/></w:r>');
      }

      for (int i = 0; i < text.length; i++) {
        final String t = text[i] ?? '';
        String closeHyperlink = '';
        String openHyperlink = '';
        if (textStyles[i] != null &&
            textStyles[i].hyperlinkTo != null &&
            textStyles[i].hyperlinkTo.isNotEmpty) {
          _packager.addHyperlink(textStyles[i].hyperlinkTo);
          closeHyperlink = '</w:hyperlink>';
          openHyperlink = '<w:hyperlink r:id="rId${_packager.rIdCount - 1}">';
        }
        _docxstring.writeAll(<String>[
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
        _addToCharCounters(t);
      }

      final String singleBreak =
          lineOrPageBreak != null && !addBreakAfterEveryItem
              ? _getLineOrPageBreak(lineOrPageBreak)
              : '';

      _docxstring.write('$singleBreak</w:p>');
      _parCount++;
    }
  }

  /// Convert inches to EMU; A4 format is 8.3 x 11.7 inches.
  int convertInchesToEMU(int inches) => inches * 914400;

  /// Convert centimeters to EMU; A4 format is 21x29.7cm.
  int convertCentimetersToEMU(int cm) => cm * 360000;

  /// Convert millimeters to EMU; A4 format is 210x297mm.
  int convertMillimetersToEMU(int mm) => mm * 36000;

  /// AddImageWithText inserts an anchor image from a file positioned at the absolute offset coordinates on the page (in EMU). [text] is optional and will wrap as determined by [wrappingAroundImage] and [textWrapping].
  /// There are a lot of options for anchor images, for example an anchor image can look like a background image by setting [behindDoc] to true.
  /// Ensure that the image is compressed to minimize the size of the docx file.
  ///
  /// The image can act like a hyperlink, by setting a destination in [hyperlinkTo]. A [description] can be attached also.
  ///
  /// The dimensions of A4:
  ///
  /// width: 210mm/21cm/8.3 inches
  ///
  /// height: 297mm/29.7cm/11.7 inches
  ///
  void addImageWithText(
    File imageFile,
    int widthEMU,
    int heightEMU,
    int horizontalOffsetEMU,
    int verticalOffsetEMU, {
    // behindDoc="0" distT="0" distB="0" distL="0" distR="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1" relativeHeight="2"
    int distT = 0,
    int distB = 0,
    int distL = 0,
    int distR = 0,
    bool simplePos = false,
    bool locked = false,
    bool layoutInCell = true,
    bool allowOverlap = true,
    bool behindDoc = false,
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
    String description = '',
    bool noChangeAspect = true,
    bool noChangeArrowheads = true,
    bool noMove = true,
    bool noResize = true,
    bool noSelect = false,
    String hyperlinkTo,
    List<String> text,
    DocxTextStyle imageTextStyle,
    List<DocxTextStyle> textStyles,
    DocxPageStyle pageStyle,
    bool doNotUseGlobalTextStyle = false,
    bool doNotUseGlobalPageStyle = true,
    LineBreak lineOrPageBreak,
    bool addBreakAfterEveryItem = false,
    bool addTab = false,
    int effectExtentL = 0,
    int effectExtentT = 0,
    int effectExtentR = 0,
    int effectExtentB = 0,
  }) {
    const int max = 27273042316900;
    final style =
        imageTextStyle ?? DocxTextStyle(textAlignment: TextAlignment.center);
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
          _docxstring.write(
              '$openParagraph$openPpr$openR<w:drawing><wp:anchor behindDoc="$behindDoc" distT="$distT" distB="$distB" distL="$distL" distR="$distR" simplePos="$simplePos" locked="$locked" layoutInCell="$layoutInCell" allowOverlap="$allowOverlap" relativeHeight="$relativeHeight"><wp:simplePos x="$simplePosX" y="$simplePosY" /><wp:positionH relativeFrom="${getValueFromEnum(horizontalPositionRelativeBase)}"><wp:posOffset>$horizontalOffsetEMU</wp:posOffset></wp:positionH><wp:positionV relativeFrom="${getValueFromEnum(verticalPositionRelativeBase)}"><wp:posOffset>$verticalOffsetEMU</wp:posOffset></wp:positionV><wp:extent cx="$widthEMU" cy="$heightEMU"/><wp:effectExtent l="$effectExtentL" t="$effectExtentT" r="$effectExtentR" b="$effectExtentB"/><wp:${getValueFromEnum(anchorImageAreaWrap)} wrapText="${getValueFromEnum(anchorImageTextWrap)}"/><wp:docPr id="$mediaIdCount" name="Image$mediaIdCount" descr="$description">$hyperlink</wp:docPr><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="$_noChangeAspect" noMove="$_noMove" noResize="$_noResize" noSelect="$_noSelect"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="$mediaIdCount" name="Image$mediaIdCount" descr="$description"></pic:cNvPr><pic:cNvPicPr><a:picLocks noChangeAspect="$_noChangeAspect" noMove="$_noMove" noResize="$_noResize" noSelect="$_noSelect" noChangeArrowheads="$_noChangeArrowheads"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId$mediaIdCount"></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr bwMode="auto"></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>$closeR$closeParagraph');
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
          addText(
            '',
            lineOrPageBreak: LineBreak(
                lineBreakType: LineBreakType.textWrapping,
                lineBreakClearLocation: LineBreakClearLocation.all),
          );
        }
      }
    }
  }

  /// AddImage inserts an inline image from a file at the current position in the buffer. AddImage is useful for stand-alone images that do not have accompanying text.
  /// If you want multiple images in the same paragraph, use AddImages. If you need text around the image, use addImageWithText. You can also use tables to create a grid of images and/or text.
  ///
  /// Ensure that the image is compressed to minimize the size of the docx file.
  ///
  /// Width and height should be provided in EMU and should be between 1 and 27273042316900.
  /// The width and height should respect the aspect ratio (width / height) of your image to prevent distortion.
  /// You can use convertMillimetersToEMU, convertCentimetersToEMU and convertInchesToEMU to calculate EMU.
  /// The constant emuWidthA4Pct and emuHeightA4Pct values can also be used.
  ///
  /// The image can act like a hyperlink, by setting [hyperlinkTo]. A [description] can be attached also.
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
    String description = '',
    bool noChangeAspect = true,
    bool noMove = true,
    bool noResize = true,
    bool noSelect = false,
    DocxTextStyle textStyle,
    bool doNotUseGlobalStyle = true,
    String hyperlinkTo,
    int effectExtentL = 0,
    int effectExtentT = 0,
    int effectExtentR = 0,
    int effectExtentB = 0,
  }) =>
      _insertInlineImage(
        imageFile,
        widthEMU,
        heightEMU,
        description: description,
        noChangeAspect: noChangeAspect,
        noMove: noMove,
        noResize: noResize,
        noSelect: noSelect,
        textStyle: textStyle,
        doNotUseGlobalStyle: doNotUseGlobalStyle,
        encloseInParagraph: true,
        hyperlinkTo: hyperlinkTo,
      );

  /// Add multiple images inside the same paragraph.
  /// All given images will be placed with the same widthEMU and heightEMU.
  /// Ensure that all images are compressed to minimize the size of the docx file.
  ///
  /// The width and height should respect the aspect ratio (width / height) of your images to prevent distortion.
  /// You can use convertMillimetersToEMU, convertCentimetersToEMU and convertInchesToEMU to calculate EMU.
  /// The constant emuWidthA4Pct and emuHeightA4Pct values can also be used.
  ///
  /// If given, the [descriptions] and [hyperlinksTo] lists for the images should have the same length as imageFiles.
  /// You can add distance between the images with [spaces] which mimic spacebar presses.
  void addImages(
    List<File> imageFiles,
    int widthEMU,
    int heightEMU, {
    List<String> descriptions,
    List<String> hyperlinksTo,
    bool noChangeAspect = true,
    bool noMove = true,
    bool noResize = true,
    bool noSelect = false,
    DocxTextStyle textStyle,
    bool doNotUseGlobalStyle = true,
    int spaces = 0,
    int effectExtentL = 0,
    int effectExtentT = 0,
    int effectExtentR = 0,
    int effectExtentB = 0,
  }) {
    final style =
        textStyle ?? DocxTextStyle(textAlignment: TextAlignment.center);

    final List<String> desc =
        descriptions != null && descriptions.length == imageFiles.length
            ? descriptions
            : List.generate(imageFiles.length, (index) => '');

    final List<String> hyperlinks =
        hyperlinksTo != null && hyperlinksTo.length == imageFiles.length
            ? hyperlinksTo
            : List.generate(imageFiles.length, (index) => '');

    final String openParagraphWithPpr =
        '<w:p>${_getParagraphStyleAsString(textStyle: style, doNotUseGlobalStyle: doNotUseGlobalStyle)}<w:r>';
    _docxstring.write(openParagraphWithPpr);
    final String addSpaces = spaces > 0
        ? '<w:t xml:space="preserve">${List<String>.generate(spaces, (index) => ' ').join()}</w:t>'
        : '';
    for (int i = 0; i < imageFiles.length; i++) {
      _insertInlineImage(
        imageFiles[i],
        widthEMU,
        heightEMU,
        description: desc[i] ?? '',
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
      );
      if (addSpaces.isNotEmpty) {
        _docxstring.write(addSpaces);
      }
    }
    _docxstring.write('</w:r></w:p>');
  }

  void _insertInlineImage(
    File imageFile,
    int widthEMU,
    int heightEMU, {
    String description = '',
    bool noChangeAspect = true,
    bool noMove = true,
    bool noResize = true,
    bool noSelect = false,
    DocxTextStyle textStyle,
    bool doNotUseGlobalStyle = true,
    bool encloseInParagraph = true,
    String hyperlinkTo,
    int effectExtentL = 0,
    int effectExtentT = 0,
    int effectExtentR = 0,
    int effectExtentB = 0,
  }) {
    const int max = 27273042316900;
    final style =
        textStyle ?? DocxTextStyle(textAlignment: TextAlignment.center);
    if (widthEMU < 1 || widthEMU > max || heightEMU < 1 || heightEMU > max) {
      return;
    }

    if (!_bufferClosed) {
      final String path = imageFile.path;
      final String suffix = path.substring(path.lastIndexOf('.') + 1);
      if (mimeTypes.containsKey(suffix)) {
        final bool saved = _packager.addImageFile(imageFile, suffix);
        final String _noChangeAspect = noChangeAspect ? '1' : '0';
        final String _noMove = noMove ? '1' : '0';
        final String _noResize = noResize ? '1' : '0';
        final String _noSelect = noSelect ? '1' : '0';

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

        if (saved) {
          _docxstring.write(
              '$openParagraph$openPpr$openR<w:drawing><wp:inline><wp:extent cx="$widthEMU" cy="$heightEMU"/><wp:effectExtent l="$effectExtentL" t="$effectExtentT" r="$effectExtentR" b="$effectExtentB"/><wp:docPr id="$mediaIdCount" name="Image$mediaIdCount" descr="$description">$hyperlink</wp:docPr><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="$_noChangeAspect" noMove="$_noMove" noResize="$_noResize" noSelect="$_noSelect"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="$mediaIdCount" name="Image$mediaIdCount" descr="$description"></pic:cNvPr><pic:cNvPicPr><a:picLocks noChangeAspect="$_noChangeAspect" noMove="$_noMove" noResize="$_noResize" noSelect="$_noSelect" noChangeArrowheads="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId$mediaIdCount"></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr bwMode="auto"></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>$closeR$closeParagraph');
        }
      }
    }
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
        final String lastSectPr =
            _getDocxPageStyleAsString(style: _globalDocxPageStyle);
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
