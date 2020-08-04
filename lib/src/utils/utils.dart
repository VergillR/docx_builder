/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

/// Valid format is 'RRGGBB' without a leading '#'
bool isValidColor(String c) =>
    c != null && c.length == 6 && RegExp(r'^[0-9a-fA-F]{6}$').hasMatch(c);

String getValueFromEnum(dynamic e) {
  final String v = e.toString();
  return v.substring(v.lastIndexOf('.') + 1);
}

String getUtcTimestampFromDateTime(DateTime d) {
  final String result = d.toUtc().toIso8601String();
  return '${result.substring(0, result.lastIndexOf('.'))}Z';
}
