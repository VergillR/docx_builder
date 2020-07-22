import 'package:docx_builder/src/styles/styles.dart';

class Shading {
  final String shadingColor;
  final ShadingPatternStyle shadingPattern;
  final String shadingPatternColor;

  Shading(
      {this.shadingColor = 'FFFFFF',
      this.shadingPattern = ShadingPatternStyle.nil,
      this.shadingPatternColor = 'FFFFFF'});
}
