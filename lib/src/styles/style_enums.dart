/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

enum Anchor { margin, page, text }

enum AnchorImageAreaWrap {
  wrapNone,
  wrapSquare,
  wrapTight,
  wrapThrough,
  wrapTopAndBottom,
}

enum AnchorImageHorizontalAlignment { left, right, center, inside, outside }

enum AnchorImageVerticalAlignment { top, bottom, center, inside, outside }

enum AnchorImageTextWrap { bothSides, left, right, largest }

enum DropCap { margin, drop, none }

/// Default footer is used when firstPageFooter and evenPageFooter are not used. EvenPageFooter and OddPageFooter are not used at the same time (only OddPageFooter) unless useEvenAndOddHeaders (despite its name, it also applies to footers) is set to true.
enum FooterType {
  evenPage,
  oddPage,
  firstPage,
}

/// Default header is used when firstPageHeader and evenPageHeader are not set. EvenPageHeader and OddPageHeader are not used at the same time (only OddPageHeader) unless useEvenAndOddHeaders was set to true in settings.xml.
enum HeaderType {
  evenPage,
  oddPage,
  firstPage,
}

enum HighlightColor {
  black,
  blue,
  cyan,
  darkBlue,
  darkCyan,
  darkGray,
  darkGreen,
  darkMagenta,
  darkRed,
  darkYellow,
  green,
  lightGray,
  magenta,
  none,
  red,
  white,
  yellow,
}

enum HorizontalPositionRelativeBase {
  margin,
  page,
  column,
  character,
  leftMargin,
  rightMargin,
  insideMargin,
  outsideMargin,
}

enum HRule { auto, atLeast, exact }

enum LineBreakClearLocation { all, left, none, right }

enum LineBreakType { column, page, textWrapping }

enum LineRule { auto, exact, atLeast }

enum NumberingList { bullet, numbered, custom }

enum NumberingSuffix { nothing, space, tab }

enum NumberFormat {
  decimal,
  upperRoman,
  lowerRoman,
  upperLetter,
  lowerLetter,
  ordinal,
  cardinalText,
  ordinalText,
  hex,
  chicago,
  ideographDigital,
  japaneseCounting,
  aiueo,
  iroha,
  decimalFullWidth,
  decimalHalfWidth,
  japaneseLegal,
  japaneseDigitalTenThousand,
  decimalEnclosedCircle,
  decimalFullWidth2,
  aiueoFullWidth,
  irohaFullWidth,
  decimalZero,
  bullet,
  ganada,
  chosung,
  decimalEnclosedFullstop,
  decimalEnclosedParen,
  decimalEnclosedCircleChinese,
  ideographEnclosedCircle,
  ideographTraditional,
  ideographZodiac,
  ideographZodiacTraditional,
  taiwaneseCounting,
  ideographLegalTraditional,
  taiwaneseCountingThousand,
  taiwaneseDigital,
  chineseCounting,
  chineseLegalSimplified,
  chineseCountingThousand,
  koreanDigital,
  koreanCounting,
  koreanLegal,
  koreanDigital2,
  vietnameseCounting,
  russianLower,
  russianUpper,
  none,
  numberInDash,
  hebrew1,
  hebrew2,
  arabicAlpha,
  arabicAbjad,
  hindiVowels,
  hindiConsonants,
  hindiNumbers,
  hindiCounting,
  thaiLetters,
  thaiNumbers,
  thaiCounting
}

enum PageBorderDisplay { allPages, firstPage, notFirstPage }

enum PageBorderSide { top, bottom, left, right }

enum PageOrientation { landscape, portrait }

enum PageNumberingFormat {
  cardinalText,
  decimal,
  decimalEnclosedCircle,
  decimalEnclosedFullstop,
  decimalEnclosedParen,
  decimalZero,
  lowerLetter,
  lowerRoman,
  none,
  ordinalText,
  upperLetter,
  upperRoman,
}

enum ParagraphBorderSide { top, bottom, left, right, between, bar }

enum ParagraphBorderStyle {
  nil,
  none,
  single,
  thick,
  double,
  dotted,
  dashed,
  dotDash,
  dotDotDash,
  triple,
  thinThickSmallGap,
  thickThinSmallGap,
  thinThickThinSmallGap,
  thinThickMediumGap,
  thickThinMediumGap,
  thinThickThinMediumGap,
  thinThickLargeGap,
  thickThinLargeGap,
  thinThickThinLargeGap,
  wave,
  doubleWave,
  dashSmallGap,
  dashDotStroked,
  threeDEmboss,
  threeDEngrave,
  outset,
  inset,
  apples,
  archedScallops,
  babyPacifier,
  babyRattle,
  balloons3Colors,
  balloonsHotAir,
  basicBlackDashes,
  basicBlackDots,
  basicBlackSquares,
  basicThinLines,
  basicWhiteDashes,
  basicWhiteDots,
  basicWhiteSquares,
  basicWideInline,
  basicWideMidline,
  basicWideOutline,
  bats,
  birds,
  birdsFlight,
  cabins,
  cakeSlice,
  candyCorn,
  celticKnotwork,
  certificateBanner,
  chainLink,
  champagneBottle,
  checkedBarBlack,
  checkedBarColor,
  checkered,
  christmasTree,
  circlesLines,
  circlesRectangles,
  classicalWave,
  clocks,
  compass,
  confetti,
  confettiGrays,
  confettiOutline,
  confettiStreamers,
  confettiWhite,
  cornerTriangles,
  couponCutoutDashes,
  couponCutoutDots,
  crazyMaze,
  creaturesButterfly,
  creaturesFish,
  creaturesInsects,
  creaturesLadyBug,
  crossStitch,
  cup,
  decoArch,
  decoArchColor,
  decoBlocks,
  diamondsGray,
  doubleD,
  doubleDiamonds,
  earth1,
  earth2,
  eclipsingSquares1,
  eclipsingSquares2,
  eggsBlack,
  fans,
  film,
  firecrackers,
  flowersBlockPrint,
  flowersDaisies,
  flowersModern1,
  flowersModern2,
  flowersPansy,
  flowersRedRose,
  flowersRoses,
  flowersTeacup,
  flowersTiny,
  gems,
  gingerbreadMan,
  gradient,
  handmade1,
  handmade2,
  heartBalloon,
  heartGray,
  hearts,
  heebieJeebies,
  holly,
  houseFunky,
  hypnotic,
  iceCreamCones,
  lightBulb,
  lightning1,
  lightning2,
  mapPins,
  mapleLeaf,
  mapleMuffins,
  marquee,
  marqueeToothed,
  moons,
  mosaic,
  musicNotes,
  northwest,
  ovals,
  packages,
  palmsBlack,
  palmsColor,
  paperClips,
  papyrus,
  partyFavor,
  partyGlass,
  pencils,
  people,
  peopleWaving,
  peopleHats,
  poinsettias,
  postageStamp,
  pumpkin1,
  pushPinNote2,
  pushPinNote1,
  pyramids,
  pyramidsAbove,
  quadrants,
  rings,
  safari,
  sawtooth,
  sawtoothGray,
  scaredCat,
  seattle,
  shadowedSquares,
  sharksTeeth,
  shorebirdTracks,
  skyrocket,
  snowflakeFancy,
  snowflakes,
  sombrero,
  southwest,
  stars,
  starsTop,
  stars3d,
  starsBlack,
  starsShadowed,
  sun,
  swirligig,
  tornPaper,
  tornPaperBlack,
  trees,
  triangleParty,
  triangles,
  tribal1,
  tribal2,
  tribal3,
  tribal4,
  tribal5,
  tribal6,
  twistedLines1,
  twistedLines2,
  vine,
  waveline,
  weavingAngles,
  weavingBraid,
  weavingRibbon,
  weavingStrips,
  whiteFlowers,
  woodwork,
  xIllusions,
  zanyTriangles,
  zigZag,
  zigZagStitch,
}

enum PreferredWidthType { auto, dxa, nil, pct }

enum RestartLineNumber { continuous, newPage, newSection }

enum SectType {
  continuous,
  evenPage,
  nextColumn,
  nextPage,
  oddPage,
}

enum SectVerticalAlign {
  both,
  bottom,
  center,
  top,
}

enum ShadingPatternStyle {
  nil,
  clear,
  solid,
  horzStripe,
  vertStripe,
  reverseDiagStripe,
  diagStripe,
  horzCross,
  diagCross,
  thinHorzStripe,
  thinVertStripe,
  thinReverseDiagStripe,
  thinDiagStripe,
  thinHorzCross,
  thinDiagCross,
  pct5,
  pct10,
  pct12,
  pct15,
  pct20,
  pct25,
  pct30,
  pct35,
  pct37,
  pct40,
  pct45,
  pct50,
  pct55,
  pct60,
  pct62,
  pct65,
  pct70,
  pct75,
  pct80,
  pct85,
  pct87,
  pct90,
  pct95,
}

enum TabLeader {
  dot,
  heavy,
  hyphen,
  middleDot,
  none,
  underscore,
}

enum TabStyle { bar, center, clear, decimal, end, num, start, left, right }

enum TableBorderSide { top, bottom, left, right, start, end, insideH, insideV }

enum TableCellBorderSide {
  top,
  bottom,
  left,
  right,
  start,
  end,
  insideH,
  insideV,
  tl2br,
  tr2bl
}

enum TableCellSide { top, bottom, left, right, start, end }

enum TableCellVerticalAlignment { bottom, center, top }

enum TableConditionalFormatting {
  firstColumn,
  firstRow,
  lastColumn,
  lastRow,
  noHBand,
  noVBand,
}

enum TableTextAlignment { left, right, center, start, end }

enum TextAlignment { left, right, center, distribute, justify, start, end }

enum TextArtStyle {
  emboss,
  imprint,
  outline,
  shadow,
}

enum TextEffect {
  blinkBackground,
  lights,
  antsBlack,
  antsRed,
  shimmer,
  sparkle,
  none,
}

enum TextFrameWrapping {
  around,
  auto,
  none,
  notBeside,
  through,
  tight,
}

enum UnderlineStyle {
  dash,
  dashDotDotHeavy,
  dashDotHeavy,
  dashedHeavy,
  dashLong,
  dashLongHeavy,
  dotDash,
  dotDotDash,
  dotted,
  dottedHeavy,
  double,
  none,
  single,
  thick,
  wave,
  wavyDouble,
  wavyHeavy,
  words,
}

enum VerticalAlignStyle { baseline, subscript, superscript }

enum VerticalPositionRelativeBase {
  margin,
  page,
  paragraph,
  line,
  topMargin,
  bottomMargin,
  insideMargin,
  outsideMargin,
}

enum VerticalTextAlignment {
  auto,
  baseline,
  bottom,
  center,
  top,
}

enum XPosAlign { center, inside, left, outside, right }

enum YPosAlign { bottom, center, inline, inside, outside, top }
