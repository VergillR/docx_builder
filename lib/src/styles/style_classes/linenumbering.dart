import '../../utils/utils.dart';
import '../style.dart';
import '../style_enums.dart';

class LineNumbering extends Style {
  final int countBy;
  final int distance;
  final RestartLineNumber restartLineNumber;
  final int start;

  LineNumbering(
      {this.countBy = 1,
      this.distance = 1440,
      this.restartLineNumber = RestartLineNumber.continuous,
      this.start = 1});

  @override
  String getXml() =>
      '<w:lnNumType w:countBy="$countBy" w:start="$start" w:restart="${getValueFromEnum(restartLineNumber)}" w:distance="$distance" />';
}
