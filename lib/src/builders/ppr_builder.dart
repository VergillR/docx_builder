import 'package:docx_builder/src/styles/style_enums.dart';
import 'package:docx_builder/src/styles/style_classes/index.dart';
import 'package:docx_builder/src/utils/utils.dart';

class Ppr {
  Ppr();

  /// getPpr returns the paragraph properties as a String.
  ///
  /// Only set properties that CHANGE the style in the current paragraph.
  /// The default style starts all values as false/disabled.
  /// rprString is obtained with Rpr.getRpr({ ...runproperties }); sectPrString with Spr.getSpr({...})
  /// Colors must be provided in "RRGGBB" format.
  static String getPpr({
    String rprString,
    String sectPrString,
    TextAlignment textAlignment,
    ParagraphBorder paragraphBorderOnAllSides,
    List<ParagraphBorder> paragraphBorders,
    PIndent paragraphIndent,
    Shading paragraphShading,
    PSpacing spacing,
    List<DocxTab> tabs,
    VerticalTextAlignment verticalTextAlignment,
    bool keepLines,
    bool keepNext,
    TextFrame textFrame,
    NumberingList numberingList,
    int numberLevelInList,
  }) {
    final StringBuffer p = StringBuffer()..write('<w:pPr>');
    if (textFrame != null) {
      p.write(textFrame.getXml());
    }
    if (numberingList != null) {
      final int numId = numberingList == NumberingList.bullet
          ? 1
          : numberingList == NumberingList.numbered ? 2 : 3;
      p.write(
          '<w:numPr><w:ilvl w:val="${numberLevelInList ?? 0}"/><w:numId w:val="$numId"/></w:numPr>');
    }
    if (keepLines != null) {
      p.write('<w:keepLines val="$keepLines" />');
    }
    if (keepNext != null) {
      p.write('<w:keepNext val="$keepNext" />');
    }
    if (textAlignment != null) {
      p.write('<w:jc w:val="${getValueFromEnum(textAlignment)}" />');
    }
    if (paragraphIndent != null && paragraphIndent.indents.isNotEmpty) {
      p.write(paragraphIndent.getXml());
    }
    if (verticalTextAlignment != null) {
      p.write(
          '<w:textAlignment w:val="${getValueFromEnum(verticalTextAlignment)}" />');
    }
    if (paragraphBorderOnAllSides != null) {
      // 'between' and 'bar' can be omitted when all borders are identical
      const List<String> sides = <String>['top', 'bottom', 'left', 'right'];
      final String borderColor = isValidColor(paragraphBorderOnAllSides.color)
          ? paragraphBorderOnAllSides.color
          : '000000';
      p.write('<w:pBdr>');
      for (int i = 0; i < sides.length; i++) {
        p.write(
            '<w:${sides[i]} w:val="${getValueFromEnum(paragraphBorderOnAllSides.pbrStyle)}" w:sz="${paragraphBorderOnAllSides.width}" w:space="${paragraphBorderOnAllSides.space}" w:color="$borderColor" w:shadow="${paragraphBorderOnAllSides.shadow}" />');
      }
      p.write('</w:pBdr>');
    } else if (paragraphBorders != null &&
        paragraphBorders.isNotEmpty &&
        paragraphBorders.length <= ParagraphBorderSide.values.length) {
      p.write('<w:pBdr>');
      for (int i = 0; i < paragraphBorders.length; i++) {
        final ParagraphBorder border = paragraphBorders[i];
        p.write(border.getXml());
      }
      p.write('</w:pBdr>');
    }
    if (paragraphShading != null) {
      p.write(paragraphShading.getXml());
    }
    if (spacing != null) {
      p.write(spacing.getXml());
    }
    if (tabs != null && tabs.isNotEmpty) {
      final StringBuffer tab = StringBuffer();
      for (int i = 0; i < tabs.length; i++) {
        final DocxTab target = tabs[i];
        tab.write(target.getXml());
      }
      p.write('<w:tabs>${tab.toString()}</w:tabs>');
    }

    if (rprString != null) {
      p.write(rprString);
    }

    if (sectPrString != null) {
      p.write(sectPrString);
    }

    p.write('</w:pPr>');
    return p.toString();
  }
}
