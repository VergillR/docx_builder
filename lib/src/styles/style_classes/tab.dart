import 'package:docx_builder/src/styles/styles.dart';

class DocxTab {
  final TabLeader leader;
  final TabStyle style;
  final int position;

  DocxTab({
    this.leader = TabLeader.none,
    this.style = TabStyle.start,
    this.position = 1,
  });
}
