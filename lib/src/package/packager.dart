import 'dart:io';

import 'package:flutter_archive/flutter_archive.dart';
import 'package:docx_builder/src/package/package_builders/index.dart';

const String cacheDocXBuilder = '_cacheDocxBuilder';
const String src = 'source';
const String dest = 'output';

class Packager {
  final Directory cacheDirectory;
  String _dirPathToDocProps;
  String _dirPathToRels;
  String _dirPathToWord;
  String _dirPathToWordMedia;
  String _dirPathToWordRels;

  Packager(this.cacheDirectory) {
    _dirPathToDocProps =
        '${cacheDirectory.path}/$src/$cacheDocXBuilder/docProps';
    _dirPathToRels = '${cacheDirectory.path}/$src/$cacheDocXBuilder/_rels';
    _dirPathToWord = '${cacheDirectory.path}/$src/$cacheDocXBuilder/word';
    _dirPathToWordMedia =
        '${cacheDirectory.path}/$src/$cacheDocXBuilder/word/media';
    _dirPathToWordRels =
        '${cacheDirectory.path}/$src/$cacheDocXBuilder/word/_rels';
    _initPackager();
  }

  void _initPackager() {
    Directory(_dirPathToDocProps).createSync(recursive: true);
    Directory(_dirPathToRels).createSync(recursive: true);
    Directory(_dirPathToWord).createSync(recursive: true);
    Directory(_dirPathToWordMedia).createSync(recursive: true);
    Directory(_dirPathToWordRels).createSync(recursive: true);
  }

  /// Bundles all files and outputs the file if successful.
  /// The file is in the cache, so user has to store or send it right away before the cache is cleared.
  Future<File> bundle(
    String documentXml, {
    int chars,
    int charsWithSpaces,
    int paragraphs,
  }) async {
    try {
      final String filename = 'D${DateTime.now().millisecondsSinceEpoch}.docx';
      final String destination =
          '${cacheDirectory.path}/$cacheDocXBuilder/$dest/$filename';
      final Directory sourceDirectory =
          Directory('${cacheDirectory.path}/$cacheDocXBuilder/$src');

      File('${cacheDirectory.path}/$cacheDocXBuilder/$src/[Content_Types].xml')
          .writeAsStringSync(ContentTypes().getContentsTypesXml());
      File('$_dirPathToDocProps/app.xml').writeAsStringSync(AppXml(
        chars: chars,
        charsWithSpaces: charsWithSpaces,
        paragraphs: paragraphs,
      ).getAppXml());
      File('$_dirPathToDocProps/core.xml')
          .writeAsStringSync(CoreXml().getCoreXml());
      File('$_dirPathToWordRels/document.xml.rels')
          .writeAsStringSync(DocumentXmlRels().getDocumentXmlRels());
      File('$_dirPathToWord/document.xml').writeAsStringSync(documentXml);
      File('$_dirPathToWord/fontTable.xml')
          .writeAsStringSync(FontTable().getFontTableXml());
      File('$_dirPathToWord/settings.xml')
          .writeAsStringSync(SettingsXml().getSettingsXml());
      File('$_dirPathToWord/styles.xml')
          .writeAsStringSync(WordStylesXml().getWordStylesXml());

      final File docxFile = File(destination);
      await ZipFile.createFromDirectory(
          sourceDir: sourceDirectory, zipFile: docxFile, recurseSubDirs: true);
      return docxFile;
    } catch (e) {
      throw StateError('DOCX BUILDER ERROR: $e');
    }
  }

  void destroyCache() {
    if (Directory('${cacheDirectory.path}/$cacheDocXBuilder').existsSync()) {
      Directory('${cacheDirectory.path}/$cacheDocXBuilder')
          .deleteSync(recursive: true);
    }
  }
}
