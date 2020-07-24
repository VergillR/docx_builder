import 'package:docx_builder/src/styles/styles.dart';
import 'package:docx_builder/src/styles/style_classes/index.dart';
import 'package:docx_builder/src/utils/utils.dart';

class Rpr {
  Rpr();

  /// getRpr returns the run properties as a String
  ///
  /// Only set properties that CHANGE the style in the current run. The default style starts all values as false/disabled.
  /// Colors must be provided in "RRGGBB" format (except for highlightColor).
  /// Note that some properties should not be active at the same time, e.g. smallCaps & caps, strike & doubleStrike.
  static String getRpr({
    bool bold,
    bool caps,
    bool doubleStrike,
    String fontColor,
    String fontName,
    String fontNameComplexScript,
    int fontSize,
    HighlightColor highlightColor,
    bool italic,
    bool rtlText,
    Shading shading,
    bool smallCaps,
    bool strike,
    TextArtStyle textArt,
    UnderlineStyle underline,
    String underlineColor,
    bool vanish,
    VerticalAlignStyle vertAlign,
    int spacing,
    int stretchOrCompressInPercentage,
    int kern,
    int fitText,
    TextEffect effect,
    bool styleAsHyperlink = false,
  }) {
    final StringBuffer r = StringBuffer()..write('<w:rPr>');
    if (styleAsHyperlink) {
      r.write('<w:rStyle w:val="Hyperlink" />');
    }
    if (fontName != null) {
      r.write(fontNameComplexScript != null
          ? '<w:rFonts w:ascii="$fontName" w:hAnsi="$fontName" w:cs="$fontNameComplexScript" />'
          : '<w:rFonts w:ascii="$fontName" w:hAnsi="$fontName" />');
    }
    if (rtlText != null) {
      r.write('<w:rtl w:val="$rtlText" />');
    }
    if (effect != null) {
      r.write('<w:effect w:val="${getValueFromEnum(effect)}" />');
    }
    if (spacing != null) {
      r.write('<w:spacing w:val="$spacing" />');
    }
    if (stretchOrCompressInPercentage != null &&
        stretchOrCompressInPercentage >= 1 &&
        stretchOrCompressInPercentage <= 600) {
      r.write('<w:w w:val="$stretchOrCompressInPercentage" />');
    }
    if (kern != null) {
      r.write('<w:kern w:val="$kern" />');
    }
    if (fitText != null) {
      r.write('<w:fitText w:val="$fitText" />');
    }
    if (shading != null) {
      final String shadingColor =
          isValidColor(shading.shadingColor) ? shading.shadingColor : 'FFFFFF';
      final String shadingPatternColor =
          isValidColor(shading.shadingPatternColor)
              ? shading.shadingPatternColor
              : 'FFFFFF';
      r.write(
          '<w:shd w:val="${getValueFromEnum(shading.shadingPattern)}" w:fill="$shadingPatternColor" w:color="$shadingColor" />');
    }
    if (highlightColor != null) {
      r.write('<w:highlight w:val="${getValueFromEnum(highlightColor)}" />');
    }
    if (bold != null) {
      r.write('<w:b w:val="$bold" /><w:bCs w:val="$bold" />');
    }
    if (italic != null) {
      r.write('<w:i w:val="$italic" /><w:iCs w:val="$italic" />');
    }
    if (strike != null) {
      r.write('<w:strike w:val="$strike" />');
    }
    if (doubleStrike != null) {
      r.write('<w:dstrike w:val="$doubleStrike" />');
    }
    if (vanish != null) {
      r.write('<w:vanish w:val="$vanish" />');
    }
    if (fontSize != null) {
      r.write(
          '<w:sz w:val="${fontSize * 2}" /><w:szCs w:val="${fontSize * 2}" />');
    }
    if (caps != null) {
      r.write('<w:caps w:val="$caps" />');
    }
    if (smallCaps != null) {
      r.write('<w:smallCaps w:val="$smallCaps" />');
    }
    if (underline != null) {
      r.write(
          '<w:u w:val="${getValueFromEnum(underline)}" w:color="${underlineColor != null && isValidColor(underlineColor) ? underlineColor : "auto"}"/>');
    }
    if (fontColor != null && isValidColor(fontColor)) {
      r.write('<w:color w:val="${fontColor.toUpperCase()}" />');
    }
    if (textArt != null) {
      r.write('<w:${getValueFromEnum(textArt)} />');
    }
    if (vertAlign != null) {
      r.write(
          '<w:vertAlign w:val="${vertAlign == VerticalAlignStyle.baseline ? 'baseline' : vertAlign == VerticalAlignStyle.subscript ? 'subscript' : 'superscript'}" />');
    }

    r.writeln('</w:rPr>');
    return r.toString();
  }
}
