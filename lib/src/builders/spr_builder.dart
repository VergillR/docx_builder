import 'package:docx_builder/src/styles/style_enums.dart';
import 'package:docx_builder/src/styles/style_classes/index.dart';
import 'package:docx_builder/src/utils/utils.dart';

class SectPr {
  SectPr();

  static String getSpr({
    List<SectColumn> cols,
    bool colsHaveSeparator = false,
    bool colsHaveEqualWidth = false,
    SectType sectType,
    int pageSzHeight, // = 16839,
    int pageSzWidth, // = 11907,
    PageOrientation pageOrientation,
    SectVerticalAlign sectVerticalAlign,
    List<PageBorder> pageBorders,
    PageBorderDisplay pageBorderDisplay,
    bool pageBorderOffsetBasedOnText = false,
    bool pageBorderIsRenderedAboveText = false,
    PageNumberingFormat pageNumberingFormat,
    int pageNumberingStart,
    LineNumbering lineNumbering,
    PageMargin pageMargin,
  }) {
    final StringBuffer s = StringBuffer()..write('<w:sectPr>');

    if (pageOrientation != null ||
        pageSzHeight != null ||
        pageSzWidth != null) {
      final szh = pageSzHeight != null ? 'w:h="$pageSzHeight"' : '';
      final szw = pageSzWidth != null ? 'w:w="$pageSzWidth"' : '';
      final String o = pageOrientation == null
          ? ''
          : pageOrientation == PageOrientation.portrait
              ? 'w:orient="portrait"'
              : 'w:orient="landscape"';
      s.write('<w:pgSz $szh $szw $o />');
    }

    if (pageMargin != null) {
      s.write(pageMargin.getXml());
    }

    if (cols != null && cols.isNotEmpty) {
      final int numCols = cols.length;
      if (colsHaveEqualWidth) {
        final int width = cols[0].spaceAfterColumn;
        s.write(
            '<w:cols w:num="$numCols" w:sep="$colsHaveSeparator" w:space="$width" />');
      } else {
        s.write(
            '<w:cols w:num="$numCols" w:sep="$colsHaveSeparator" w:space="720" w:equalWidth="false">');
        for (int i = 0; i < numCols; i++) {
          s.write(cols[i].getXml());
        }
        s.write('</w:cols>');
      }
    }
    if (sectType != null) {
      s.write('<w:type w:val="${getValueFromEnum(sectType)}" />');
    }

    if (sectVerticalAlign != null) {
      s.write('<w:vAlign w:val="${getValueFromEnum(sectVerticalAlign)}"/>');
    }

    if (pageBorders != null &&
        pageBorders.isNotEmpty &&
        pageBorders.length <= PageBorderSide.values.length) {
      final String pageBorderOffset =
          pageBorderOffsetBasedOnText ? "text" : "page";
      final String pageBorderZOrder =
          pageBorderIsRenderedAboveText ? "front" : "back";
      final String display = pageBorderDisplay == null
          ? ""
          : pageBorderDisplay == PageBorderDisplay.allPages
              ? 'display="allPages"'
              : pageBorderDisplay == PageBorderDisplay.firstPage
                  ? 'display="firstPage"'
                  : 'display="notFirstPage"';

      s.write(
          '<w:pgBorders w:offsetFrom="$pageBorderOffset" zOrder="$pageBorderZOrder" $display>');
      for (int j = 0; j < pageBorders.length; j++) {
        final PageBorder border = pageBorders[j];
        s.write(border.getXml());
      }
      s.write('</w:pgBorders>');
    }

    final String start =
        pageNumberingStart == null ? '' : 'w:start="$pageNumberingStart" ';
    final String format = pageNumberingFormat == null
        ? ''
        : 'w:fmt="${getValueFromEnum(pageNumberingFormat)}"';
    if (start.isNotEmpty || format.isNotEmpty) {
      s.write('<w:pgNumType $start$format/>');
    }

    if (lineNumbering != null) {
      s.write(lineNumbering.getXml());
    }

    s.write('</w:sectPr>');
    return s.toString();
  }
}
