import '../style.dart';
import '../style_enums.dart';

class PSpacing extends Style {
  final List<int> spaces;
  final LineRule lineRule;
  final bool beforeAutospacing;
  final bool afterAutospacing;

  PSpacing({
    int after = 0,
    int afterLines = 0,
    int before = 0,
    int beforeLines = 0,
    int line = 0,
    this.lineRule = LineRule.atLeast,
    this.beforeAutospacing = false,
    this.afterAutospacing = false,
  }) : spaces = <int>[after, afterLines, before, beforeLines, line];

  @override
  String getXml() {
    const List<String> sides = <String>[
      'after',
      'afterLines',
      'before',
      'beforeLines',
      'line',
    ];

    final StringBuffer spbuffer = StringBuffer();
    for (int i = 0; i < spaces.length; i++) {
      final int val = spaces[i];
      if (val > 0) {
        spbuffer.write('w:${sides[i]}="$val" ');
      }
    }

    return '<w:spacing ${spbuffer.toString()} w:beforeAutospacing="$beforeAutospacing" w:afterAutospacing="$afterAutospacing" w:lineRule="${lineRule == LineRule.auto ? "auto" : lineRule == LineRule.atLeast ? "atLeast" : "exactly"}" />';
  }
}
