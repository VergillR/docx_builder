import 'package:docx_builder/src/utils/constants/constants.dart';

class CoreXml {
  CoreXml();

  String getCoreXml({
    String documentTitle = '',
    String documentSubject = '',
    String documentDescription = '',
    String documentCreator = '',
  }) =>
      '$docPropsCoreXml<dc:title>$documentTitle</dc:title><dc:subject>$documentSubject</dc:subject><dc:description>$documentDescription</dc:description><dc:creator>$documentCreator</dc:creator></cp:coreProperties>';
}
