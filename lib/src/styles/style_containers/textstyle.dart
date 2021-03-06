/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import '../style_classes/index.dart';
import '../style_enums.dart';

class TextStyle {
  // rpr (Run properties: text styling)
  final String hyperlinkTo;
  final bool bold;
  final bool caps;
  final bool doubleStrike;
  final String fontColor;
  final String fontName;
  final String fontNameComplexScript;
  final int fontSize;
  final HighlightColor highlightColor;
  final bool italic;
  final bool rtlText;
  final Shading shading;
  final bool smallCaps;
  final bool strike;
  final TextArtStyle textArt;
  final UnderlineStyle underline;
  final String underlineColor;
  final bool vanish;
  final VerticalAlignStyle vertAlign;
  final int textSpacing;
  final int stretchOrCompressInPercentage;
  final int kern;
  final int fitText;
  final TextEffect textEffect;
  // ppr (Paragraph properties: paragraph styling)
  final TextAlignment textAlignment;
  final ParagraphBorder paragraphBorderOnAllSides;
  final List<ParagraphBorder> paragraphBorders;
  final PIndent paragraphIndent;
  final Shading paragraphShading;
  final PSpacing paragraphSpacing;
  final List<DocxTab> tabs;
  final VerticalTextAlignment verticalTextAlignment;
  final bool keepLines;
  final bool keepNext;
  final TextFrame textFrame;
  final NumberingList numberingList;
  final int numberLevelInList;

  TextStyle({
    // rpr (Run properties: text styling)
    this.hyperlinkTo,
    this.bold,
    this.caps,
    this.doubleStrike,
    this.fontColor,
    this.fontName,
    this.fontNameComplexScript,
    this.fontSize,
    this.highlightColor,
    this.italic,
    this.rtlText,
    this.shading,
    this.smallCaps,
    this.strike,
    this.textArt,
    this.underline,
    this.underlineColor,
    this.vanish,
    this.vertAlign,
    this.textSpacing,
    this.stretchOrCompressInPercentage,
    this.kern,
    this.fitText,
    this.textEffect,
    // ppr (Paragraph properties: paragraph styling)
    this.textAlignment,
    this.paragraphBorderOnAllSides,
    this.paragraphBorders,
    this.paragraphIndent,
    this.paragraphShading,
    this.paragraphSpacing,
    this.tabs,
    this.verticalTextAlignment,
    this.keepLines,
    this.keepNext,
    this.textFrame,
    this.numberingList,
    this.numberLevelInList,
  });
}
