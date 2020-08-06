/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import '../style_classes/index.dart';
import '../style_enums.dart';

/// PageStyle is the set of Sections Properties of the document that change the page's appearance.
/// Most (if not all) can be omitted, allowing the word processor to use its default styling and rules.
///
/// [pageNumberingFormat] should only be set once in a section; [pageNumberingStart] can usually be omitted, but if it is set, it indicates at which number the page counter resets whenever a new section is encountered.
class PageStyle {
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
  final PageNumberingFormat footnoteNumberingFormat;
  final PageNumberingFormat endnoteNumberingFormat;

  PageStyle({
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
    this.footnoteNumberingFormat,
    this.endnoteNumberingFormat,
  });

  PageStyle getDefaultPageStyle() => PageStyle(
        pageSzHeight: 16839,
        pageSzWidth: 11907,
      );
}
