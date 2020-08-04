/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import 'package:docx_builder/src/utils/constants/constants.dart';

class ContentTypes {
  ContentTypes();

  // Replace designated target with default and mime types of included media to contentTypes.
  String getContentsTypesXml(
          Set<String> defaultRefs, Set<String> overrideRefs) =>
      contentTypes.replaceFirst(
          reservedTarget, defaultRefs.join() + overrideRefs.join());
}
