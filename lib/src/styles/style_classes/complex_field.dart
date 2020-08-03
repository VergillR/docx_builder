import '../style.dart';

class ComplexField extends Style {
  final String instructions;
  final bool includeSeparate;
  final String hyperlinkText;

  ComplexField(
      {this.instructions,
      this.includeSeparate = true,
      this.hyperlinkText = ''});

  @override
  String getXml() {
    final String sep =
        includeSeparate ? '<w:fldChar w:fldCharType="separate"/>' : '';
    return '<w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">$instructions</w:instrText></w:r><w:r>$sep</w:r><w:r>${hyperlinkText.isNotEmpty ? "<w:rPr><w:rStyle w:val=\"Hyperlink\" /></w:rPr>" : ""}<w:t>$hyperlinkText</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/>';
  }
}
