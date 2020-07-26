import 'package:docx_builder/src/styles/style_enums.dart';

abstract class Header {
  String rId;
  final HeaderType headerType;
  String taggedText;

  Header({
    this.rId,
    this.headerType,
    this.taggedText,
  });
}

class EvenPageHeader extends Header {
  EvenPageHeader({
    String taggedText,
    String rId,
  }) : super(headerType: HeaderType.evenPage, taggedText: taggedText, rId: rId);
}

/// Also called default header.
class OddPageHeader extends Header {
  OddPageHeader({String taggedText, String rId})
      : super(headerType: HeaderType.oddPage, taggedText: taggedText, rId: rId);
}

/// Also called titlePageHeader.
class FirstPageHeader extends Header {
  FirstPageHeader({String taggedText, String rId})
      : super(
            headerType: HeaderType.firstPage, taggedText: taggedText, rId: rId);
}
