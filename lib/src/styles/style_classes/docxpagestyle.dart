import '../style_enums.dart';
import 'index.dart';

class DocxPageStyle {
  final List<SectColumn> cols;
  final bool colsHaveSeparator;
  final bool colsHaveEqualWidth;
  final SectType sectType;
  final int pageSzHeight;
  final int pageSzWidth;
  final PageOrientation pageOrientation;
  final SectVerticalAlign sectVerticalAlign;
  final List<PageBorder> pageBorders;
  final PageBorderDisplay pageBorderDisplay;
  final bool pageBorderOffsetBasedOnText;
  final bool pageBorderIsRenderedAboveText;
  final PageNumberingFormat pageNumberingFormat;
  final int pageNumberingStart;
  final LineNumbering lineNumbering;
  final PageMargin pageMargin;

  DocxPageStyle({
    this.cols,
    this.colsHaveSeparator = false,
    this.colsHaveEqualWidth = false,
    this.sectType,
    this.pageSzHeight,
    this.pageSzWidth,
    this.pageOrientation,
    this.sectVerticalAlign,
    this.pageBorders,
    this.pageBorderDisplay,
    this.pageBorderOffsetBasedOnText = false,
    this.pageBorderIsRenderedAboveText = false,
    this.pageNumberingFormat,
    this.pageNumberingStart,
    this.lineNumbering,
    this.pageMargin,
  });

  DocxPageStyle getDefaultPageStyle() => DocxPageStyle(
        pageNumberingFormat: PageNumberingFormat.decimal,
        pageOrientation: PageOrientation.portrait,
        pageSzHeight: 16839,
        pageSzWidth: 11907,
      );
}
