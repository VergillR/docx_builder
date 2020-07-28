import '../../utils/utils.dart';
import '../style_enums.dart';

/// Width and height are in twips (twentiehs of a point). TextFrame is a paragraph property.
class TextFrame {
  final bool anchorLock;
  final DropCap dropCap;
  final HRule hRule;
  final int height;
  final Anchor hAnchor;
  final int hSpace;
  final int linesHeightDropCap;
  final Anchor vAnchor;
  final int vSpace;
  final int width;
  final TextFrameWrapping textFrameWrapping;
  final int xPosition;
  final int yPosition;
  final XPosAlign xPosAlign;
  final YPosAlign yPosAlign;

  TextFrame({
    this.anchorLock,
    this.dropCap,
    this.hRule,
    this.height,
    this.hAnchor,
    this.hSpace,
    this.linesHeightDropCap = 1,
    this.vAnchor,
    this.vSpace,
    this.width,
    this.textFrameWrapping,
    this.xPosition,
    this.yPosition,
    this.xPosAlign,
    this.yPosAlign,
  });

  String getFramePr() {
    final String aLock =
        anchorLock != null ? 'w:anchorLock="$anchorLock" ' : '';
    final String dCap =
        dropCap != null ? 'w:dropCap="${getValueFromEnum(dropCap)}" ' : '';
    final String hR =
        hRule != null ? 'w:hRule="${getValueFromEnum(hRule)}" ' : '';
    final String h = height != null ? 'w:h="$height" ' : '';
    final String hA =
        hAnchor != null ? 'w:hAnchor="${getValueFromEnum(hAnchor)}" ' : '';
    final String hS = hSpace != null ? 'w:hSpace="$hSpace" ' : '';
    final String lines = 'w:lines="$linesHeightDropCap" ';
    final String vA =
        vAnchor != null ? 'w:vAnchor="${getValueFromEnum(vAnchor)}" ' : '';
    final String vS = vSpace != null ? 'w:vSpace="$vSpace" ' : '';
    final String w = width != null ? 'w:w="$width" ' : '';
    final String wrapping = textFrameWrapping != null
        ? 'w:wrap="${getValueFromEnum(textFrameWrapping)}" '
        : '';
    final String xpos = xPosition != null ? 'w:x="$xPosition" ' : '';
    final String ypos = yPosition != null ? 'w:y="$yPosition" ' : '';
    final String xposA =
        xPosAlign != null ? 'w:xAlign="${getValueFromEnum(xPosAlign)}" ' : '';
    final String yposA =
        yPosAlign != null ? 'w:yAlign="${getValueFromEnum(yPosAlign)}" ' : '';
    return '<w:framePr $w$h$wrapping$aLock$dCap$hR$hA$hS$lines$vA$vS$xpos$ypos$xposA$yposA/>';
  }
}
