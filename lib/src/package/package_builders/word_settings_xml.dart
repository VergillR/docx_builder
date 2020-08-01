import 'package:docx_builder/src/utils/constants/constants.dart';

class SettingsXml {
  SettingsXml();

  String getSettingsXml({bool useEvenHeaders = false}) => useEvenHeaders
      ? wordSettingsXml.replaceFirst(
          '</w:settings>', '<w:evenAndOddHeaders/></w:settings>')
      : wordSettingsXml;
}
