import '../../utils/utils.dart';
import '../style.dart';
import '../style_enums.dart';

/// LineBreakClearLocation only has meaning if LineBreakType is set to 'textWrapping'
class LineBreak extends Style {
  final LineBreakType lineBreakType;
  final LineBreakClearLocation lineBreakClearLocation;

  LineBreak({
    this.lineBreakType,
    this.lineBreakClearLocation,
  });

  @override
  String getXml() =>
      '<w:r><w:br w:type="${lineBreakType != null ? getValueFromEnum(lineBreakType) : "textWrapping"}" w:clear="${lineBreakClearLocation != null ? getValueFromEnum(lineBreakClearLocation) : "none"}" /></w:r>';
}
