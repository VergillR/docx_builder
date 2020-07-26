import 'package:docx_builder/src/styles/style_enums.dart';
import 'package:docx_builder/src/styles/style_classes/index.dart';
import 'package:docx_builder/src/utils/utils.dart';

// abstract class Reference {
//   final String id;
//   final ReferenceType referenceType;

//   Reference({this.id, this.referenceType});
// }

// class HeaderReference extends Reference {
//   HeaderReference({String id, ReferenceType referenceType})
//       : super(id: id, referenceType: referenceType);
// }

// class FooterReference extends Reference {
//   FooterReference({String id, ReferenceType referenceType})
//       : super(id: id, referenceType: referenceType);
// }

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

    // final String o = pageOrientation == null
    //     ? ''
    //     : pageOrientation == PageOrientation.portrait
    //         ? 'w:orient="portrait"'
    //         : 'w:orient="landscape"';
    // s.write('<w:pgSz w:h="$pageSzHeight w:w="$pageSzWidth $o />');

    // final PageMargin pg = pageMargin ?? PageMargin();
    // s.write(
    //     'w:pgMar w:header="${pg.header}" w:footer="${pg.footer}" w:gutter="${pg.gutter}" w:left="${pg.left}" w:right="${pg.right}" w:top="${pg.top}" w:bottom="${pg.bottom}" />');

    if (pageMargin != null) {
      s.write(
          'w:pgMar w:header="${pageMargin.header}" w:footer="${pageMargin.footer}" w:gutter="${pageMargin.gutter}" w:left="${pageMargin.left}" w:right="${pageMargin.right}" w:top="${pageMargin.top}" w:bottom="${pageMargin.bottom}" />');
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
          s.write(
              'w:col w:w="${cols[i].w}" w:space="${cols[i].spaceAfterColumn}" />');
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

    if (pageBorders != null && pageBorders.isNotEmpty) {
      const List<String> sides = <String>[
        'top',
        'bottom',
        'left',
        'right',
      ];
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
        final String borderColor =
            isValidColor(border.color) ? border.color : '000000';
        s.write(
            '<w:${sides[border.pageBorderSide.index]} w:val="${getValueFromEnum(border.pbrStyle)}" w:color="$borderColor" w:sz="${border.size}" w:space="${border.space}" w:shadow="${border.shadow}" />');
      }
      s.write('</w:pgBorders>');
    }

    final String start =
        pageNumberingStart == null ? '' : 'w:fmt="$pageNumberingStart"';
    final String format = pageNumberingFormat == null
        ? ''
        : 'w:fmt="${getValueFromEnum(pageNumberingFormat)}"';
    if (start.isNotEmpty || format.isNotEmpty) {
      s.write('<w:pgNumType $start $format />');
    }

    if (lineNumbering != null) {
      s.write(
          '<w:lnNumType w:countBy="${lineNumbering.countBy}" w:start="${lineNumbering.start}" w:restart="${getValueFromEnum(lineNumbering.restartLineNumber)}" w:distance="${lineNumbering.distance}" />');
    }

    s.writeln('</w:sectPr>');
    return s.toString();
  }
}
