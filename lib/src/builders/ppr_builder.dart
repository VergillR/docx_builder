import 'package:docx_builder/src/styles/styles.dart';
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
    PBorder paragraphBorderOnAllSides,
    List<PBorder> paragraphBorders,
    PIndent paragraphIndent,
    Shading paragraphShading,
    PSpacing spacing,
    List<Tab> tabs,
    VerticalTextAlignment vAlign,
    bool keepLines,
    bool keepNext,
  }) {
    final StringBuffer p = StringBuffer()..write('<w:pPr>');
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
      const List<String> indents = [
        'left',
        'start',
        'right',
        'end',
        'hanging',
        'firstLine'
      ];
      final StringBuffer ind = StringBuffer();
      for (int i = 0; i < paragraphIndent.indents.length; i++) {
        final int val = paragraphIndent.indents[i];
        if (val > 0) {
          ind.write('w:${indents[i]}="$val" ');
        }
      }
      p.write('<w:ind ${ind.toString()}/>');
    }

    if (vAlign != null) {
      p.write('<w:textAlignment w:val="${getValueFromEnum(vAlign)}" />');
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
        paragraphBorders.length < 7) {
      const List<String> sides = <String>[
        'top',
        'bottom',
        'left',
        'right',
        'between',
        'bar',
      ];
      p.write('<w:pBdr>');
      for (int i = 0; i < paragraphBorders.length; i++) {
        final PBorder border = paragraphBorders[i];
        final String borderColor =
            isValidColor(border.color) ? border.color : '000000';
        p.write(
            '<w:${sides[border.borderSide.index]} w:val="${getValueFromEnum(paragraphBorderOnAllSides.pbrStyle)}" w:sz="${border.width}" w:space="${border.space}" w:color="$borderColor" w:shadow="${border.shadow}" />');
      }
      p.write('</w:pBdr>');
    }
    if (paragraphShading != null) {
      final String shadingColor = isValidColor(paragraphShading.shadingColor)
          ? paragraphShading.shadingColor
          : 'FFFFFF';
      final String shadingPatternColor =
          isValidColor(paragraphShading.shadingPatternColor)
              ? paragraphShading.shadingPatternColor
              : 'FFFFFF';
      final String sp = paragraphShading.shadingPattern.toString();
      final String sh = sp.substring(sp.lastIndexOf('.') + 1);
      p.write(
          '<w:shd w:val="$sh" w:fill="$shadingPatternColor" w:color="$shadingColor" />');
    }
    if (spacing != null) {
      const List<String> sides = <String>[
        'after',
        'afterLines',
        'before',
        'beforeLines',
        'line',
      ];

      final StringBuffer spbuffer = StringBuffer();
      for (int i = 0; i < spacing.spaces.length; i++) {
        final int val = spacing.spaces[i];
        if (val > 0) {
          spbuffer.write('w:${sides[i]}="$val" ');
        }
      }
      p.write(
          '<w:spacing ${spbuffer.toString()} w:beforeAutospacing="${spacing.beforeAutospacing}" w:afterAutospacing="${spacing.afterAutospacing}" w:lineRule="${spacing.lineRule == LineRule.auto ? "auto" : spacing.lineRule == LineRule.atLeast ? "atLeast" : "exactly"}" />');
    }
    if (tabs != null && tabs.isNotEmpty) {
      const List<String> tLeader = [
        'dot',
        'heavy',
        'hyphen',
        'middleDot',
        'none',
        'underscore'
      ];
      const List<String> tStyle = [
        'bar',
        'center',
        'clear',
        'decimal',
        'end',
        'num',
        'start'
      ];
      final StringBuffer tab = StringBuffer();
      for (int i = 0; i < tabs.length; i++) {
        final Tab target = tabs[i];
        tab.write(
            '<w:tab w:val="${tStyle[target.style.index]}" w:leader="${tLeader[target.leader.index]}" pos="${target.position}" />');
      }
      p.write('<w:tabs>${tab.toString()}</w:tabs>');
    }

    if (rprString != null) {
      p.write(rprString);
    }

    if (sectPrString != null) {
      p.write(sectPrString);
    }

    p.writeln('</w:pPr>');
    return p.toString();
  }
}
