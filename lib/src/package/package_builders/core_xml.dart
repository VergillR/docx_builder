/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import 'package:docx_builder/src/utils/constants/constants.dart';

class CoreXml {
  CoreXml();

  String getCoreXml({
    String documentTitle = '',
    String documentSubject = '',
    String documentDescription = '',
    String documentCreator = '',
  }) =>
      '$docPropsCoreXml<dc:title>$documentTitle</dc:title><dc:subject>$documentSubject</dc:subject><dc:description>$documentDescription</dc:description><dc:creator>$documentCreator</dc:creator></cp:coreProperties>';
}
