import '../../utils/utils.dart';
import '../style.dart';
import '../style_enums.dart';

class PageBorder extends Style {
  final PageBorderSide pageBorderSide;
  final int size;
  final int space;
  final String color;
  final ParagraphBorderStyle pbrStyle;
  final bool shadow;

  PageBorder({
    this.pageBorderSide = PageBorderSide.bottom,
    this.size = 8,
    this.space = 24,
    this.color = '000000',
    this.pbrStyle = ParagraphBorderStyle.single,
    this.shadow = false,
  });

  @override
  String getXml() {
    final String borderColor = isValidColor(color) ? color : '000000';
    return '<w:${getValueFromEnum(pageBorderSide)} w:val="${getValueFromEnum(pbrStyle)}" w:color="$borderColor" w:sz="$size" w:space="$space" w:shadow="$shadow" />';
  }
}
