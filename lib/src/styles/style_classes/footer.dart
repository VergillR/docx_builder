import 'package:docx_builder/src/styles/style_enums.dart';

abstract class Footer {
  String rId;
  final FooterType footerType;
  String taggedText;

  Footer({this.rId, this.footerType, this.taggedText});
}

class EvenPageFooter extends Footer {
  EvenPageFooter({String taggedText, String rId})
      : super(
            footerType: FooterType.evenPage, taggedText: taggedText, rId: rId);
}

/// Also called default Footer.
class OddPageFooter extends Footer {
  OddPageFooter({String taggedText, String rId})
      : super(footerType: FooterType.oddPage, taggedText: taggedText, rId: rId);
}

/// Also called titlePageFooter.
class FirstPageFooter extends Footer {
  FirstPageFooter({String taggedText, String rId})
      : super(
            footerType: FooterType.firstPage, taggedText: taggedText, rId: rId);
}
