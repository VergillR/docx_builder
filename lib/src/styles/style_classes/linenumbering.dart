import '../style_enums.dart';

class LineNumbering {
  final int countBy;
  final int distance;
  final RestartLineNumber restartLineNumber;
  final int start;

  LineNumbering(
      {this.countBy = 1,
      this.distance = 1440,
      this.restartLineNumber = RestartLineNumber.continuous,
      this.start = 1});
}
