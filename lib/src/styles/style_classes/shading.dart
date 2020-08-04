/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import 'package:docx_builder/src/styles/style_enums.dart';
import '../../utils/utils.dart';
import '../style.dart';

class Shading extends Style {
  final String shadingColor;
  final ShadingPatternStyle shadingPattern;
  final String shadingPatternColor;

  Shading(
      {this.shadingColor = 'FFFFFF',
      this.shadingPattern = ShadingPatternStyle.nil,
      this.shadingPatternColor = 'FFFFFF'});

  @override
  String getXml() {
    final String shadingCol =
        isValidColor(shadingColor) ? shadingColor : 'FFFFFF';
    final String shadingPatternCol =
        isValidColor(shadingPatternColor) ? shadingPatternColor : 'FFFFFF';
    return '<w:shd w:val="${getValueFromEnum(shadingPattern ?? ShadingPatternStyle.nil)}" w:fill="$shadingPatternCol" w:color="$shadingCol" />';
  }
}
