# Change log

## master (unreleased)
### New features
* Add ability to configure units of measurements
* Add support of `line_3d`, `bar3DChart`, `pie3DChart` charts
* Add parsing text direction in table cell
* Refactor parsing page size of document
* Refactor parsing page margins of docx
* Refactor parsing columns data of docx
* Some minor RuboCop refactor
* Add parsing `fldSimple` inside hyperlinks
* Add support of `wp14:pctPosHOffset` and `wp14:pctPosVOffset`
* Add support of parsing `w:noBreakHyphen`, `w:tab` in PargraphRun
* Add support of Table Row Properties - Height
* Add variable to store original file path in all parsers

### Fixes
* Fix parsing shape in paragraph run

## 0.1.0
* Initial release of `ooxml_parser` gem