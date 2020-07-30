import '../style.dart';
import './index.dart';

class TableRow extends Style {
  final TableRowProperties tableRowProperties;
  List<TableCell> tableCells;

  TableRow({
    this.tableRowProperties,
    this.tableCells,
  });

  @override
  String getXml() {
    final StringBuffer s = StringBuffer()..write('<w:tr>');

    if (tableRowProperties != null) {
      s.write(tableRowProperties.getXml());
    }
    if (tableCells != null && tableCells.isNotEmpty) {
      for (int i = 0; i < tableCells.length; i++) {
        if (tableCells[i] != null) {
          s.write(tableCells[i].getXml());
        } else {
          s.write('<w:tc></w:tc>');
        }
      }
    }

    s.write('</w:tr>');
    return s.toString();
  }
}
