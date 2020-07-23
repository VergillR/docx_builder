import 'package:docx_builder/src/styles/styles.dart';

/// LineBreakClearLocation only has meaning if LineBreakType is set to 'textWrapping'
class LineBreak {
  final LineBreakType lineBreakType;
  final LineBreakClearLocation lineBreakClearLocation;

  LineBreak({
    this.lineBreakType,
    this.lineBreakClearLocation,
  });
}
