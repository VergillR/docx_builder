/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import '../style.dart';

class SectColumn extends Style {
  final int w;
  final int spaceAfterColumn;

  SectColumn({this.w = 1000, this.spaceAfterColumn = 720});

  @override
  String getXml() => 'w:col w:w="$w" w:space="$spaceAfterColumn" />';
}
