import '../style.dart';

class SectColumn extends Style {
  final int w;
  final int spaceAfterColumn;

  SectColumn({this.w = 1000, this.spaceAfterColumn = 720});

  @override
  String getXml() => 'w:col w:w="$w" w:space="$spaceAfterColumn" />';
}
