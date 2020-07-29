import '../../utils/utils.dart';
import '../style.dart';
import '../style_enums.dart';

class FloatingTable extends Style {
  final Anchor horzAnchor;
  final Anchor vertAnchor;
  final int absXPos;
  final int absYPos;
  final XPosAlign relXPos;
  final YPosAlign relYPos;
  final int bottomFromText;
  final int topFromText;
  final int leftFromText;
  final int rightFromText;

  FloatingTable({
    this.horzAnchor,
    this.vertAnchor,
    this.absXPos,
    this.absYPos,
    this.relXPos,
    this.relYPos,
    this.bottomFromText,
    this.topFromText,
    this.leftFromText,
    this.rightFromText,
  });

  @override
  String getXml() {
    final String hAnchor = horzAnchor != null
        ? 'w:horzAnchor="${getValueFromEnum(horzAnchor)}" '
        : '';
    final String vAnchor = vertAnchor != null
        ? 'w:vertAnchor="${getValueFromEnum(vertAnchor)}" '
        : '';
    final String absX = absXPos != null ? 'w:tblpX="$absXPos" ' : '';
    final String absY = absYPos != null ? 'w:tblpY="$absYPos" ' : '';
    final String relX = relXPos != null ? 'w:tblpXSpec="$relXPos" ' : '';
    final String relY = relYPos != null ? 'w:tblpYSpec="$relYPos" ' : '';

    final String ll =
        leftFromText != null ? 'w:leftFromText="$leftFromText" ' : '';
    final String rr =
        rightFromText != null ? 'w:rightFromText="$rightFromText" ' : '';
    final String tt =
        topFromText != null ? 'w:topFromText="$topFromText" ' : '';
    final String bb =
        bottomFromText != null ? 'w:bottomFromText="$bottomFromText" ' : '';

    return '<w:tblpPr $hAnchor$vAnchor$absX$absY$relX$relY$ll$rr$tt$bb/>';
  }
}
