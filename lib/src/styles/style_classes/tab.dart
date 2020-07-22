import 'package:docx_builder/src/styles/styles.dart';

class Tab {
  final TabLeader leader;
  final TabStyle style;
  final int position;

  Tab({
    this.leader = TabLeader.none,
    this.style = TabStyle.start,
    this.position = 1,
  });
}
