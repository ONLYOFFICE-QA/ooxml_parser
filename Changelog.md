# Change log

## master (unreleased)
### New features
* Ability to set units of measurement to each value, not to all via config
* Add twips (same as dxa), one_eighth_point unit of measurement
* Add configuration to set accuracy of fraction part of digit
* Add parsing of Table Properties and Table Style Properties in Document Styles
* Fully support of Windows platform
* Add parsing of Shape Adjust Values List
* Add parsing Relative Sizes of shape
* Add parsing math formula run properties (and redone some math model for it)
* Add parsing math formula argument size property
* Add a whole lot new properties to parse in charts
* Remove usage of Linux `file` command in favor of `ruby-filemagick` gem. Better cross-platform support
* Use ruby method to create tmp folder instead of calling Linux methods
* Add method `OoxmlParser::Parser.parse` to parse any OOXML format with auto recognition
* Add storing color scheme name in color class
* Add parsing of properties of table - inside horizontal, vertical borders
* Add support of parsing more text directions in cell
* Add support of `show_category_name` and `show_series_name` to chart label properties
* Add method to 3 base formats to check if it contain any user data `#with_data?`
* Correct parsing of paragraph properties and run properties
* Paragraphs now have correct parsing of borders
* Run properties - language, position property
* Parsing run shade
* Parsing of Columns spacing
* Redone parsing `gridSpan` and `vMerge`
* Add ability to get style on which style is based
* Parsing `TableStyleColumnBandSize`, `TableStyleRowBandSize`, `TableLayout`, table cell spacing for TableProperties
* Correct parsing `Margins` for `tcPr`
* OoxmlSize support a whole lot formats
* Correct handling unsupported OoxmlSize format
* `DocumentStlye` parse `CellProperties`
* Clean way to parse `ParagraphSpacing`
* `ParagraphTab` now stored in `Tabs` in `ParagraphProperties`
* Add `OoxmlSize` support of `one_240th_cm` and use in `Spacing`
* Add `OoxmlSize` support of `spacing_point` and use in `Spacing`
* Add `OoxmlSize` support of `one_100th_point` and use in `RunProperties#spacing`
* Add parsing `PresetColor` to `GradientStop`
* Add Parsing `NumberingLevel#suffix`
* Add parsing `CellProperties#no_wrap`
* Add parsing `TableStyleProperties#table_properties`
* Add parsing `TableStyleProperties#paragraph_properties`
* Add parsing `ParagraphProperties#justification`
* Add parsing `TableRowProperties#table_header` (fix #264)
* Add parsing `XlsxColumnProperties#custom_width`, `XlsxColumnProperties#best_fit`

### Fixes
* Fix parsing document style id - it can be string, not only digit
* Fix misplaced `dxa` and `emu` units of measurements and also fix calculation `dxa` unit
* Redone parsing of `nary` in formulas
* Fix parsing gradient color linear values
* Fix parsing Footnote and Endnote reference in runs
* Fix calculating position offset values, distance from text in different units of measurements
* Fix parsing table style in text box
* Fix problem with parsing absolute file path in Windows 
* Parse `keep_next` in paragraph properties
* `TransformEffect`, `BordersProperties#size` in correct `OoxmlSize` unit
* `Indents`, `TableProperties#table_indent`, `ParagraphProperties#margin_left`, 
`ParagraphProperties#margin_right`, `ParagraphProperties#indent`, `DocxShapeLine#width`, 
`TextOutline#width`, `Outline#width`, `TableCellLine#width`, `XlsxDrawingPositionParameters` use `OoxmlSize`
* `TableMargins`, `TablePosition`, `FrameProperties`, `CellProperties#able_cell_width`, `TableRow#height` use `OoxmlSize`
* `TableProperties`, `TableCellProperties` use `Shade` class
* `ParagraphMargins` parse size in correct units
* `ParagraphProperties` parse `PageProperties`
* `ParagraphProperties` parse `contextual_spacing`
* Hanging indent is now 0 by default, instead of nil
* Correct parsing of `Drawing` of any type, not just `TwoCellAnchor`
* `Outline` default width in `OoxmlSize`
* `DocxShapeLine` correct zero if `nofill`
* Fix error for `DocxPicture#with_data?`
* `Worksheet#with_data?` recognize custom columns
* `Slide#with_data?` recognize custom background (Fix #256)
* `RunSpacing#value` is in OoxmlSize
* `Shade` should be able to set all argument via constructor

### Refactor
* Redone parsing images - store in structure, instead of copying file to filesystem
* Move parsing Columns inside class
* Move parsing TableGrid inside class
* Move parsing DocxShapeLinePath inside class
* Totally redone parsing of Numbering
* Parsing DocxParagraphRun properties
* Move some slide methods to helper
* Remove `OoxmlShift` - use `OoxmlCoordinates` instead
* Rename `Size` class to `PageSize` and `RunSize` to `Size`
* Simplify code for parsing `CellProperties`, `PageProperties`, 
`DocxParagraph`, `Bookmark`, `TableLook`, `ChartAxisTitle`, 
`NonVisualProperties`, `ShapePlaceholder`, `RunProperties`, 
`TableProperties`, `FrameProperties`, `Background`, `Transition`,
`TransitionProperties`, `XlsxRow`, `XlsxColumnProperties`, `CellStyle`,
`DisplayLabelsProperties`, `ParagraphProperties`, 
`ChartAxis`, `ChartLegend`, `SizeRelativeHorizontal`,
`SizeRelativeVertical`, `OOXMLShapeBodyProperties`, `Tile`,
`NaryLimitLocation`, `MultilevelType`, `CellProperties`
* Replace usage `Alignment.parse` on `OOXMLDocumentObject#value_to_symbol` method
* Refactor `Color.from_int16` to `Color#parse_hex_string` method
* Remove unused `Color#parse_int16_string`
* Remove parsing `TableProperties#right_to_left`, since it was totally wrong
* Move `table_cell_spacing` to `TableRowProperties` and use OoxmlSize
* Remove class `ParagraphStyle`
* Remove useless attribute `OOXMLDocumentObject.namespace_perfix`
* Simplify `XlsxRow` parsing.
* `FontStyle` class cleanup
* Remove unused class `TableStyleElement`
* `Border` class cleanup

## 0.1.2 (2016-06-07)

### New features
* Correct `==` method for DocxStructure.
* Correctly handle and warn if docx file do not contain `docProps/app.xml`
* Extract GradientColor parsing to Class
* Correct warn if document is password protected
* Correct warn if document links to lost image
* Add parsing Table relationships

### Fixes
#### `DocxParser`
* Fix crash while parsing `page_borders`
* Fix parsing nil borders
* Fix parsing document background and background image
* Fix parsing TextBox lists in Shape
* Fix parsing OldDocxShape image path
* Fix parsing table merge data
* Fix parsing ParagraphRun footnote-endnote
* Fix parsing `OldDocxGroup` `wrap` nil value
* Fix parsing BorderProperties space, size is nil
* Fix parsing default cell properties - space, color and shade
* Fix parsing relation target NULL
* Fix parsing `anchor_lock` in frame properties
* Fix parsing DocxBlip without properties
* Fix parsing Default Styles for Table
* Fix parsing Table Border style

#### `XlsxParser`
* Fix parsing default underline style for cell
* Fix parsing TextField in Paragraph
* Fix parsing Chart data if there is no `numRef`
* Fix parsing non-defined font name in style
* Fix parsing document theme without name
* Fix parsing nil height of row
* Fix parsing fill color nil

#### `PptxParser`
* Fix parsing Condition event and delay nil values
* Fix parsing Transition Direction, Orientation nil
* Fix parsing Sound transaction
* Fix parsing slide background fill rectangle stretching
* Fix parsing animation effect without transition
* Fix parsing click highlight in links
* Fix parsing image fill without embeded image

## 0.1.1 (2016-05-17)

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