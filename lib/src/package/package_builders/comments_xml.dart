/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import 'package:docx_builder/src/utils/constants/constants.dart';

class CommentsXml {
  CommentsXml();

  String getCommentsXml({Map<int, String> comments}) {
    final StringBuffer c = StringBuffer();
    final List<String> vals = comments.values.toList();
    for (int i = 0; i < vals.length; i++) {
      c.write(vals[i]);
    }
    return '$commentsXml${c.toString()}</w:comments>';
  }
}
