# Change log

## master (unreleased)
### New features
* Ability to set units of measurement to each value, not to all via config
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
* Run properties - language property

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

### Refactor
* Redone parsing images - store in structure, instead of copying file to filesystem
* Move parsing Columns inside class
* Move parsing TableGrid inside class
* Move parsing DocxShapeLinePath inside class
* Totally redone parsing of Numbering
* Parsing DocxParagraphRun properties
* Move some slide methods to helper
* Remove `OoxmlShift` - use `OoxmlCoordinates` instead

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