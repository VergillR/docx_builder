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
