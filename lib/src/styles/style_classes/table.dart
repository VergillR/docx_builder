import 'package:meta/meta.dart';

import 'package:docx_builder/src/styles/style_classes/index.dart';
import 'package:docx_builder/src/styles/style.dart';

class Table extends Style {
  List<int> gridColumnWidths;
  TableProperties tableProperties;
  List<TableRow> tableRows;

  Table({
    this.gridColumnWidths,
    this.tableProperties,
    this.tableRows,
  });

  @override
  String getXml() {
    if (gridColumnWidths == null ||
        gridColumnWidths.isEmpty ||
        tableRows == null ||
        tableRows.isEmpty) {
      return '';
    }
    final StringBuffer t = StringBuffer()..write('<w:tbl><w:tblGrid>');
    for (int i = 0; i < gridColumnWidths.length; i++) {
      t.write('<w:gridCol w:w="${gridColumnWidths[i]}"/>');
    }
    t.write('</w:tblGrid>');

    t.write(tableProperties.getXml());

    if (tableRows != null) {
      for (int i = 0; i < tableRows.length; i++) {
        t.write(tableRows[i].getXml());
      }
    }

    t.write('</w:tbl>');
    return t.toString();
  }

  /// Initialize or reset the table.
  ///
  /// [tableRows] contains the rows which hold the table cells carrying the actual content of the table.
  ///
  /// [gridColumnWidths] should hold widths for the maximum amount of columns, e.g. for 5 columns the list can contain the values 1600, 1000, 800, 800, 800. If the designed table has varying widths for its cells, the number of given widths should equal that of the highest number of cells possible in one row.
  ///
  /// [tableProperties] determine the global properties of the table; some properties can be overridden by row or cell properties.
  void setTable({
    @required List<int> gridColumnWidths,
    @required List<TableRow> tableRows,
    TableProperties tableProperties,
  }) {
    this.gridColumnWidths = gridColumnWidths;
    this.tableRows = tableRows;
    this.tableProperties = tableProperties;
  }
}
