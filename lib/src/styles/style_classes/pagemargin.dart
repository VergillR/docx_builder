import '../style.dart';

class PageMargin extends Style {
  final int bottom;
  final int footer;
  final int gutter;
  final int header;
  final int left;
  final int right;
  final int top;

  PageMargin({
    this.bottom = 1417,
    this.footer = 708,
    this.gutter = 0,
    this.header = 708,
    this.left = 1417,
    this.right = 1417,
    this.top = 1417,
  });

  @override
  String getXml() =>
      'w:pgMar w:header="$header" w:footer="$footer" w:gutter="$gutter" w:left="$left" w:right="$right" w:top="$top" w:bottom="$bottom" />';
}
