/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import '../../utils/utils.dart';
import '../style.dart';
import '../style_enums.dart';

class LineNumbering extends Style {
  final int countBy;
  final int distance;
  final RestartLineNumber restartLineNumber;
  final int start;

  LineNumbering(
      {this.countBy = 1,
      this.distance = 1440,
      this.restartLineNumber = RestartLineNumber.continuous,
      this.start = 1});

  @override
  String getXml() =>
      '<w:lnNumType w:countBy="$countBy" w:start="$start" w:restart="${getValueFromEnum(restartLineNumber)}" w:distance="$distance" />';
}
