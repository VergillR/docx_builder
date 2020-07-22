import 'package:docx_builder/src/styles/styles.dart';

class PSpacing {
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
}
