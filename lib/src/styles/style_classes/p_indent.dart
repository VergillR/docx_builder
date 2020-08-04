/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import '../style.dart';

class PIndent extends Style {
  List<int> indents;

  PIndent({
    int indentLeft = 0,
    int indentStart = 0,
    int indentRight = 0,
    int indentEnd = 0,
    int indentHanging = 0,
    int indentFirstLine = 0,
  }) : indents = <int>[
          indentLeft,
          indentStart,
          indentRight,
          indentEnd,
          indentHanging,
          indentFirstLine
        ];

  @override
  String getXml() {
    const List<String> inds = [
      'left',
      'start',
      'right',
      'end',
      'hanging',
      'firstLine'
    ];
    final StringBuffer s = StringBuffer();
    s.write('<w:ind ');
    for (int i = 0; i < indents.length; i++) {
      final int val = indents[i];
      if (val > 0) {
        s.write('w:${inds[i]}="$val" ');
      }
    }
    s.write('/>');
    return s.toString();
  }
}
