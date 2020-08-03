import '../../utils/utils.dart';
import '../style.dart';
import '../style_enums.dart';

class DocxTab extends Style {
  final TabLeader leader;
  final TabStyle style;
  final int position;

  DocxTab({
    this.leader = TabLeader.none,
    this.style = TabStyle.start,
    this.position = 1,
  });

  @override
  String getXml() =>
      '<w:tab w:val="${getValueFromEnum(style)}" w:leader="${getValueFromEnum(leader)}" w:pos="$position" />';
}
