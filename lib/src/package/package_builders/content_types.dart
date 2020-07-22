import 'package:docx_builder/src/utils/constants/constants.dart';
// import 'package:docx_builder/src/utils/constants/mimetypes.dart';

class ContentTypes {
  Set<String> addedMediaTypes = <String>{};

  ContentTypes();

  // TODO: Add mime types of included media to contentTypes.
  String getContentsTypesXml() => contentTypes;
}
