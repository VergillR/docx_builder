import '../styles.dart';
import 'index.dart';

class DocxTextStyle {
  // rpr (Run properties: text styling)
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
  final PBorder paragraphBorderOnAllSides;
  final List<PBorder> paragraphBorders;
  final PIndent paragraphIndent;
  final Shading paragraphShading;
  final PSpacing paragraphSpacing;
  final List<Tab> tabs;
  final VerticalTextAlignment vAlign;
  final bool keepLines;
  final bool keepNext;

  DocxTextStyle({
    // rpr (Run properties: text styling)
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
    this.vAlign,
    this.keepLines,
    this.keepNext,
  });
}
