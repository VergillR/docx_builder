class PIndent {
  List<int> indents;

  PIndent({
    int indentLeft = 0,
    int indentStart = 0,
    int indentRight = 0,
    int indentEnd = 0,
    int indentHanging = 0,
    int indentFirstLine = 0,
  }) : indents = <int>[
          indentLeft,
          indentStart,
          indentRight,
          indentEnd,
          indentHanging,
          indentFirstLine
        ];
}
