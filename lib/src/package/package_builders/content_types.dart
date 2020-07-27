import 'package:docx_builder/src/utils/constants/constants.dart';

class ContentTypes {
  ContentTypes();

  // Replace designated target with default and mime types of included media to contentTypes.
  String getContentsTypesXml(
          Set<String> defaultRefs, Set<String> overrideRefs) =>
      contentTypes.replaceFirst(
          reservedTarget, defaultRefs.join() + overrideRefs.join());
}
