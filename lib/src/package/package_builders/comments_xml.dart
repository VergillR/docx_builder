import 'package:docx_builder/src/utils/constants/constants.dart';

class CommentsXml {
  CommentsXml();

  String getCommentsXml({Map<int, String> comments}) {
    final StringBuffer c = StringBuffer();
    final List<String> vals = comments.values.toList();
    for (int i = 0; i < vals.length; i++) {
      c.write(vals[i]);
    }
    return '$commentsXml${c.toString()}</w:comments>';
  }
}
