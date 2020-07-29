import './index.dart';
import '../style_enums.dart';
import '../../builders/index.dart';

/// TableCell can contain an entire p (paragraph) or tbl (table).
class TableCell {
  final String cellId;
  final TableCellProperties tableCellProperties;

  TableCell({
    this.cellId,
    this.tableCellProperties,
  });
}
