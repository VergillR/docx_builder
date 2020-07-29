import '../style.dart';

import './index.dart';

/// TableCell can contain content which holds an entire p (paragraph with runs and text) or tbl (table).
class TableCell extends Style {
  final String cellId;
  final TableCellProperties tableCellProperties;
  final String content;

  TableCell({
    this.cellId,
    this.tableCellProperties,
    this.content,
  });

  @override
  String getXml() {
    final String id = cellId != null ? ' w:id="$cellId"' : '';
    final StringBuffer tc = StringBuffer()..write('<w:tc$id>');
    if (tableCellProperties != null) {
      tc.write(tableCellProperties.getXml());
    }
    tc.write('${content ?? '<w:p><w:r></w:r></w:p>'}</w:tc>');
    return tc.toString();
  }
}
