/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import '../../utils/utils.dart';
import '../style.dart';
import '../style_enums.dart';

class DocxTab extends Style {
  final TabLeader leader;
  final TabStyle style;
  final int position;

  DocxTab({
    this.leader = TabLeader.none,
    this.style = TabStyle.start,
    this.position = 1,
  });

  @override
  String getXml() =>
      '<w:tab w:val="${getValueFromEnum(style)}" w:leader="${getValueFromEnum(leader)}" w:pos="$position" />';
}
