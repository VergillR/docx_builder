import 'dart:io';

import 'package:docx_builder/src/utils/constants/constants.dart';
import 'package:docx_builder/src/utils/constants/mimetypes.dart';
import 'package:flutter_archive/flutter_archive.dart';
import 'package:docx_builder/src/package/package_builders/index.dart';

const String cacheDocXBuilder = '_cacheDocxBuilder';
const String src = 'source';
const String dest = 'output';
const int startIdCount = 4;

class Packager {
  final Directory cacheDirectory;
  String _dirPathToDocProps;
  String _dirPathToRels;
  String _dirPathToWord;
  String _dirPathToWordMedia;
  String _dirPathToWordRels;

  int _rIdCount = startIdCount;
  int get rIdCount => _rIdCount;

  final Map<String, String> _references = <String, String>{};

  final Set<String> _contentDefaultRefs = <String>{};
  final Set<String> _contentOverrideRefs = <String>{};

  bool _addEvenAndOddHeadersInSettings = false;

  Packager(this.cacheDirectory) {
    _dirPathToDocProps =
        '${cacheDirectory.path}/$cacheDocXBuilder/$src/docProps';
    _dirPathToRels = '${cacheDirectory.path}/$cacheDocXBuilder/$src/_rels';
    _dirPathToWord = '${cacheDirectory.path}/$cacheDocXBuilder/$src/word';
    _dirPathToWordMedia =
        '${cacheDirectory.path}/$cacheDocXBuilder/$src/word/media';
    _dirPathToWordRels =
        '${cacheDirectory.path}/$cacheDocXBuilder/$src/word/_rels';
    _initPackager();
  }

  void _initPackager() {
    Directory(_dirPathToDocProps).createSync(recursive: true);
    Directory(_dirPathToRels).createSync(recursive: true);
    Directory(_dirPathToWord).createSync(recursive: true);
    Directory(_dirPathToWordMedia).createSync(recursive: true);
    Directory(_dirPathToWordRels).createSync(recursive: true);
  }

  bool addHeaderOrFooter(int counter, String contents,
      {bool isHeader = true, bool evenPage = false}) {
    try {
      final String type = isHeader ? 'header' : 'footer';
      final String fileName = '$type$counter.xml';
      final String fullPathFile = '$_dirPathToWord/$fileName';
      if (evenPage) {
        // even page header/footer was created, so it needs to be declared in Settings.xml
        _addEvenAndOddHeadersInSettings = true;
      }
      final String totalContent = isHeader
          ? '$headerXml$contents</w:hdr>'
          : '$footerXml$contents</w:ftr>';
      File(fullPathFile).writeAsStringSync(totalContent);
      _contentOverrideRefs.add(
          '<Override PartName="/word/$fileName" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.$type+xml"/>');
      _references['rId${_rIdCount++}'] =
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/$type" Target="$fileName"';
      return true;
    } catch (e) {
      return false;
    }
  }

  void addHyperlink(String target) => _references['rId${_rIdCount++}'] =
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="$target" TargetMode="External"';

  bool addImageFile(File file, String suffix) {
    try {
      final String fileName = 'image$_rIdCount.$suffix';
      final String fullPathFile = '$_dirPathToWordMedia/$fileName';
      file.copySync(fullPathFile);
      if (mimeTypes[suffix] != null) {
        _contentDefaultRefs.add(
            '<Default Extension="$suffix" ContentType="${mimeTypes[suffix]}"/>');
      }
      _references['rId${_rIdCount++}'] =
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/$fileName"';
      return true;
    } catch (e) {
      return false;
    }
  }

  /// Bundles all files and outputs the file if successful.
  /// The file is in the cache, so user has to store or send it right away before the cache is cleared.
  Future<File> bundle(
    String documentXml, {
    // int chars,
    // int charsWithSpaces,
    // int paragraphs,
    String documentTitle,
    String documentSubject,
    String documentDescription,
    String documentCreator,
  }) async {
    try {
      final String filename = 'D${DateTime.now().millisecondsSinceEpoch}.docx';
      final String destination =
          '${cacheDirectory.path}/$cacheDocXBuilder/$dest/$filename';
      Directory('${cacheDirectory.path}/$cacheDocXBuilder/$dest')
          .createSync(recursive: false);
      final Directory sourceDirectory =
          Directory('${cacheDirectory.path}/$cacheDocXBuilder/$src');

      File('${cacheDirectory.path}/$cacheDocXBuilder/$src/[Content_Types].xml')
          .writeAsStringSync(ContentTypes()
              .getContentsTypesXml(_contentDefaultRefs, _contentOverrideRefs));
      File('$_dirPathToDocProps/app.xml').writeAsStringSync(AppXml(
              // chars: chars,
              // charsWithSpaces: charsWithSpaces,
              // paragraphs: paragraphs,
              )
          .getAppXml());
      File('$_dirPathToRels/.rels').writeAsStringSync(DotRels().getDotRels());
      File('$_dirPathToDocProps/core.xml').writeAsStringSync(
        CoreXml().getCoreXml(
          documentTitle: documentTitle ?? '',
          documentSubject: documentSubject ?? '',
          documentDescription: documentDescription ?? '',
          documentCreator: documentCreator ?? '',
        ),
      );
      File('$_dirPathToWordRels/document.xml.rels')
          .writeAsStringSync(DocumentXmlRels().getDocumentXmlRels(_references));
      File('$_dirPathToWord/document.xml').writeAsStringSync(documentXml);
      File('$_dirPathToWord/fontTable.xml')
          .writeAsStringSync(FontTable().getFontTableXml());
      File('$_dirPathToWord/settings.xml').writeAsStringSync(SettingsXml()
          .getSettingsXml(useEvenHeaders: _addEvenAndOddHeadersInSettings));
      File('$_dirPathToWord/styles.xml')
          .writeAsStringSync(WordStylesXml().getWordStylesXml());

      final File docxFile = File(destination);
      docxFile.createSync(recursive: false);
      await ZipFile.createFromDirectory(
        sourceDir: sourceDirectory,
        zipFile: docxFile,
        recurseSubDirs: true,
        includeBaseDirectory: false,
      );
      return docxFile;
    } catch (e) {
      throw StateError('DOCX BUILDER ERROR: $e');
    }
  }

  void destroyCache() {
    _rIdCount = startIdCount;
    _references.clear();
    _contentDefaultRefs.clear();
    _contentOverrideRefs.clear();
    _addEvenAndOddHeadersInSettings = false;
    if (Directory('${cacheDirectory.path}/$cacheDocXBuilder').existsSync()) {
      Directory('${cacheDirectory.path}/$cacheDocXBuilder')
          .deleteSync(recursive: true);
    }
  }
}
