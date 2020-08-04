/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import '../style.dart';

import './index.dart';

/// TableCell can contain content which holds an entire p (paragraph with runs and text) or tbl (table).
/// Content must be in formatted correctly in OOXML.
class TableCell extends Style {
  final String cellId;
  final TableCellProperties tableCellProperties;
  String xmlContent;

  TableCell({
    this.cellId,
    this.tableCellProperties,
    this.xmlContent,
  });

  @override
  String getXml() {
    final String id = cellId != null ? ' w:id="$cellId"' : '';
    final StringBuffer tc = StringBuffer()..write('<w:tc$id>');
    if (tableCellProperties != null) {
      tc.write(tableCellProperties.getXml());
    }
    tc.write('${xmlContent ?? '<w:p><w:r></w:r></w:p>'}</w:tc>');
    return tc.toString();
  }
}
