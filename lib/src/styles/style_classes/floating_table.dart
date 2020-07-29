import '../style_enums.dart';

class FloatingTable {
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
}
