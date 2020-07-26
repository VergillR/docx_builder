import 'package:docx_builder/src/utils/constants/constants.dart';

class CoreXml {
  CoreXml();

  String getCoreXml(
      {String documentTitle,
      String documentSubject,
      String documentDescription,
      String documentCreator}) {
    // ignore: prefer_final_locals
    String a = docPropsCoreXml;
    if (documentTitle.isNotEmpty) {
      a.replaceFirst('<dc:title>', '<dc:title>$documentTitle');
    }
    if (documentSubject.isNotEmpty) {
      a.replaceFirst('<dc:subject>', '<dc:subject>$documentSubject');
    }
    if (documentDescription.isNotEmpty) {
      a.replaceFirst(
          '<dc:description>', '<dc:description>$documentDescription');
    }
    if (documentCreator.isNotEmpty) {
      a.replaceFirst('<dc:creator>', '<dc:creator>$documentCreator');
    }
    return a;
  }
}
