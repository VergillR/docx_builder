import 'package:docx_builder/src/styles/style_enums.dart';
import 'package:docx_builder/src/styles/style_classes/index.dart';
import 'package:docx_builder/src/utils/utils.dart';

class Table {
  final int cols;
  final int rows;
  final List<List<String>> grid;

  Table({
    this.cols = 1,
    this.rows = 1,
  }) : grid = List.generate(
            cols, (index) => List<String>.generate(rows, (index) => ''));

  String getTbl({
    List<int> gridColumnWidths,
    TableProperties tableProperties,
    List<TableRow> tableRows,
  }) {
    if (gridColumnWidths == null ||
        gridColumnWidths.length != cols ||
        tableRows == null ||
        tableRows.length != rows) {
      return '';
    }
    final StringBuffer t = StringBuffer()..write('<w:tbl>');
    t.write('<w:tblGrid>');
    for (int i = 0; i < gridColumnWidths.length; i++) {
      t.write('<w:gridCol w:w="${gridColumnWidths[i]}"/>');
    }
    t.write('</w:tblGrid>');

    t.write('</w:tbl>');
    return t.toString();
  }
}
