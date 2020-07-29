import 'package:docx_builder/src/utils/constants/constants.dart';

class AppXml {
  // final int chars;
  // final int charsWithSpaces;
  // final int paragraphs;

  // AppXml({this.chars = 0, this.charsWithSpaces = 0, this.paragraphs = 0});

  AppXml();

  String getAppXml() => docPropsAppXml;

  // String getAppXml() => docPropsAppXml.replaceAll('</Properties>',
  //     '<Characters>$chars</Characters><CharactersWithSpaces>$charsWithSpaces</CharactersWithSpaces><Paragraphs>$paragraphs</Paragraphs></Properties>');
}
