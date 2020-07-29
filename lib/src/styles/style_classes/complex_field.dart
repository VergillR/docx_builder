import '../style.dart';

class ComplexField extends Style {
  final String instructions;
  final bool includeSeparate;

  ComplexField({this.instructions, this.includeSeparate = false});

  @override
  String getXml() {
    final String sep =
        includeSeparate ? '<w:fldChar w:fldCharType="separate"/>' : '';
    return '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">$instructions</w:instrText></w:r><w:r>$sep</w:r><w:r><w:t></w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
  }
}
