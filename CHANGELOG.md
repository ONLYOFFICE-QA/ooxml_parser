# Change log

## master (unreleased)

## 0.39.0 (2025-06-18)

### Fixes

* Fix `ParagraphBorders#border_visual_type`
  for cases when document contains empty `pBdr` tag
* Fix `rubocop-1.75.8` cop `Style/RedundantParentheses`
* Fix `rubocop-1.75.8` cop `Style/EmptyStringInsideInterpolation`
* Fix `rubocop-1.72.2` cop `Lint/UselessConstantScoping`
* Fix `rubocop-performance-1.24.0` cop `Performance/ChainArrayAllocation`
* Run `rubocop` in CI through `bundle exec`

## 0.38.0 (2025-02-17)

### New Features

* Add `ruby-3.3` to CI
* Add `ruby-3.4` to CI
* Add `dependabot` check for `GitHub Actions`
* Add support of `truffleruby`
* Add `example` folder with example how to verify some files
* Add parsing `Comment` parameters `author`, `date_string`, `initials`
* Add parsing `CommentExtended#parent_paragraph_id`
* Add parsing `UserProtectedRanges`

### Changes

* **BREAKING** Drop `ruby-2.7` support, since it's EOLed
* **BREAKING** Drop `ruby-3.0` support, since it's EOLed
* Move long running rubies CI to nightly runs
* Ability to run nightly CI by trigger button
* Optimize performance for parsing sheet name coordinates
* Code changes after update to `rubocop-1.63.1`
* Fix `rubocop-1.64` cop `Style/SuperArguments` warnings.
* Fix rubocop configuration to be compatible with v3 of `rubocop-rspec`
* Fix `rubocop-1.65.0` cop `Gemspec/AddRuntimeDependency`
* Increase coverage for some edge cases for future `ruby-3.4`
* New optional param for `NumberingProperties#numbering_level_current`
* Code changes after update to `rubocop-1.68.0`

### Fixes

* Monkey-patch `File::SHARE_DELETE` to be compatible with `truffleruby-24.1.1`
* Fix `Coordinates#column_number` for multi-chart strings

## 0.37.1 (2023-07-06)

* Add parsing `PivotField#name`

## 0.37.0 (2023-06-20)

* Add parsing `PivotTableDefinition#data_fields`
* Add parsing `PivotTableDefinition#page_fields`
* Add parsing `PivotTableDefinition#row_fields`
* Add parsing `PivotTableDefinition#column_fields`

## 0.36.1 (2023-05-22)

### Fixes

* Fix merging hyperlink runs for `DocxParagraph`

### Changes

* Store all development dependencies in Gemfile
  (According to rubocop `Gemspec/DevelopmentDependencies`)

## 0.36.0 (2023-05-10)

### New Features

* New `Color#within_delta?` method for comparing colors
* New `PageProperties#section_break` method
* Cache file `Relationships` to reduce memory usage

### Changes

* `OOXMLDocumentObject#boolean_attribute_value` will raise `ArgumentError` if
  attribute value is not boolean/1/0
* `Note#note_base_xpath` will raise `NameError` for unknown note type
* Minor changes to `spec` folder structure
* `TextBox` now is instance of `OOXMLDocumentObject`
* Cleanup `Runs`-related code in `DocxParagraph`
* Simplify `Spacing` parsing
* Add `TimeNodeList` class for parsing nodes
* Add `XlsxColumns` class for parsing nodes

### Fixes

* Fix `CommonTiming` was not an instance of `OOXMLDocumentObject
* Simplify parsing of `Chart`
* Speedup `Color#parse_hex_string` by not using regexp
* Remove all `DocumentStructure` class variables
* Add `SeriesText#string` alias to `SeriesText#reference`
* Optimize `rspec` by not parsing same files several times

## 0.35.0 (2023-04-20)

### Changes

* Drop support of `ruby-2.6` since it's EOLed
* Add jruby-9.4 support
* Fix parsing borders with `none` style
* Add parsing `SeriesText` with number reference

## 0.34.2 (2022-11-30)

### Fixes

* Fix parsing default_run_style
* Fix parsing indents

## 0.34.1 (2022-11-10)

### Fixes

* Do not crash if default paragraph style have no run properties
* Fix parsing boolean `ValuedChild` if there is no `val` attribute
* Fix early exit from parsing default run style
* Fix `NumberingProperties#ilvl` default value
* Fix problem with parsing default spacing
* Fix problem with parsing default indent

## 0.34.0 (2022-11-09)

### New Features

* Add support of `jruby-9.3`

### Changes

* Fix incorrect attributes type for `Location`
* Move `rubocop` check in CI to `linting` config
* Refactor `Colo#parse_hex_string` for better performance
* **Breaking** `Underline#style` now is always symbol

### Fixes

* Fix file parsing outside of `it` in `rspec`
* Fix parsing `FormProperties`
* Fix parsing `FormTextProperties`

## 0.33.0 (2022-10-25)

### New Features

* Add parsing `ComboBox`
* Add parsing `DropdownList`
* Add parsing `CheckBox`
* Add parsing `FormProperties`
* Add parsing `FormTextProperties`
* Add `OoxmlColor#theme_tint`

### Changes

* Refactor parsing Run color (remove strange line)
* Remove `spec` folder from coverage reports
* Several minor improvements to reduce memory usage

## 0.32.0 (2022-10-03)

### New Features

* Add parsing `WorkbookProperties`
* Add parsing `NumberingProperties#image`

## 0.31.0 (2022-09-27)

### New Features

* Add parsing `RunProperties#ligatures`

### Fixes

* Fix regexp for parsing `Coordinates#parser_coordinates_range`

### Changes

* Use `OOXMLDocumentObject#parse_xml` in all cases

## 0.30.0 (2022-09-09)

### New Features

* New class `OoxmlFile` for base operation with file

### Changes

* `PresentationTheme#parse` now an instance method
* Remove all class variables from `OOXMLDocumentObject`

### Fixes

* Fix comparing `OOXMLDocumentObject` if element has `Nokogiri::XML::Element` attribute

## 0.29.0 (2022-09-05)

### Changes

* Replace usage of `StyleParameter` to `DocumentStyle`
* Optimize parsing Style data for paragraph
* `DocumentStructure#parse` now is an instance method

### Fixes

* Fix parsing links in files with OLE spreadsheets

## 0.28.0 (2022-09-05)

### New Features

* Add parsing of xlsx contained in `OleObject`

## 0.27.0 (2022-08-31)

### New Features

* Parsing of `XlsxRow#style_index`
* New `Worksheet#rows_raw` field
* Add parsing `Chart#relationships`
* Add parsing `Chart#style`

### Changes

* Simplify code for `Worksheet` parsing.

## 0.26.0 (2022-08-22)

### New Features

* New `XlsxColumnProperties#width_raw` field
* New `XlsxRow#cells_raw` field

### Changes

* `spec` speedup by decreasing files size
* Rename `XlsxColumnProperties#from/to` to `XlsxColumnProperties#min/max` for
  better code consistency. Deprecated old methods
* Redone parsing `XlsxRow#cells` for better readability
* Minor refactoring in parsing `DocxWrapDrawing#wrap_text`
* Remove old comments from `RubyMine` code inspection

### Fixes

* Fix `Color#to_hex` if color is not initialized

## 0.25.0 (2022-07-28)

### New Features

* New CI task to check that documentation is correct.
* Correct parsing of `DocxParagraphRun#instruction_text`

### Changes

* Minor change in `Presentation` parsing to remove deprecated `Nokogiri`
  method `XML::Reader#attribute_nodes`
* Removed unused `DocxParagraph#bookmark_start` and
  `DocxParagraph#bookmark_end` attributes

## 0.24.0 (2022-06-10)

### New Features

* Add parsing `sparklines` in xlsx

### Changes

* Minor refactor of project directory structure

## 0.23.0 (2022-05-01)

### New Features

* Add `yamllint` check in CI
* New `StyleMatrixReference` class, replacing similar ones

### Changes

* Drop `ruby-2.5` support since `nokogiri` v1.13.0 dropped it
* Actualize `nodejs` version in CI
* Check `dependabot` at 8:00 Moscow time daily
* Changes from `rubocop-rspec` update to 2.9.0
* Remove `ruby-filemagick` dependency
* Remove deprecated `Coordinates#get_column_number` method
* Refactor `Coordinates.parse_coordinates_from_string`
  to `Coordinates#parse_string`

## 0.22.0 (2022-01-10)

### New Features

* Add `ruby-3.1` in CI
* Add parsing `workbookProtection` in xlsx
* Add parsing password protected files
* Add parsing `sheetProtection` in xlsx
* Add parsing `protectedRange` in xlsx
* Add parsing `Xf#protection`
* Add parsing `DocxShape#locks_text`
* Add parsing `XlsxDrawing#client_data`

### Changes

* Remove `ruby-2.5` from CI since it's EOLed

## 0.21.0 (2021-12-20)

### New Features

* Add parsing `sheets` in xlsx
* Add parsing `sheetView` attributes in xlsx
* Add parsing `selection` in xlsx

## 0.20.0 (2021-11-18)

### Fixes

* Fix crash if `Hyperlink` for slide has no `id`

## 0.19.0 (2021-11-17)

### Fixes

* Fix `ParargraphSpacing#line` parsing if different order of attributes

### Changes

* Require `mfa` for releasing gem

## 0.18.1 (2021-11-11)

### Fixes

* Fix failure if `Sound` class has no `name`

## 0.18.0 (2021-11-11)

### New Features

* Add `CodeQL` check in CI
* Add parsing `Presentation#slide_master`
* Add parsing `Presentation#slide_layouts`
* Fail if `nokogiri` found any error in file

### Changes

* Minor refactoring of `RunProperties`.
  Decrease Rubocop metrics

## 0.17.0 (2021-09-15)

### New Features

* Add parsing `ConditionalFormattingRule#above_average`
* Add parsing `ConditionalFormattingRule#equal_average`

## 0.16.0 (2021-09-14)

### New Features

* Add parsing `ConditionalFormattingRule#bottom`
* Add parsing `ConditionalFormattingRule#time_period`

### Changes

* Minor refactoring in parsing Paragraph spacing
* Minor style fixes from `rubocop` v1.21.0

## 0.15.0 (2021-09-13)

### New Features

* Parsing `Conditional Formatting` in xlsx

## 0.14.2 (2021-09-03)

### Fixes

* Fix problem with no `require 'uri'` (Thanks [@545ch4](https://github.com/545ch4))

### Changes

* Minor style fixes from `rubocop` v1.19.0

## 0.14.1 (2021-08-03)

### Fixes

* Fix `DocxParagraph#with_data?` failure on paragraph without properties

## 0.14.0 (2021-07-05)

### New Features

* Parsing `Header and Footer` in xlsx
* Parsing `Defined Names` in xlsx

## 0.13.0 (2021-05-25)

### New Features

* Parsing `Data Validation` in xlsx

### Changes

* Add `pkg` to `.gitignore`

### Removal

* Remove deprecated `DocxParagraph#frame_properties`

## 0.12.2 (2021-02-18)

### New Features

* Add parsing `ParagraphProperties#paragraph_style_ref`

## 0.12.1 (2021-02-17)

### New Features

* Add parsing `RunProperties#shade`

## 0.12.0 (2021-02-17)

### New Features

* Add parsing `ParagraphProperties#shade`

### Changes

* Add `ruby-3.0` in CI
* Deprecated `DocxParagraph#background_color`

## 0.11.0 (2020-12-10)

### New Features

* Add support of parsing `PivotCacheDefinition#cacheFields`
* Add support of parsing `CacheFields#cache_field`
* Add support of parsing `CacheField#shared_item`
* Add support of parsing `PivotTableDefinition` properties
* Add support of parsing `PivotTableDefinition#location`
* Add support of parsing `PivotTableDefinition#pivot_fields`
* Add support of parsing `PivotField#item`
* Add support of parsing `PivotTableDefinition#column_items`
* Add support of parsing `PivotTableDefinition#row_items`
* Add support of parsing `PivotTableDefinition#style_info`

### Changes

* Changes from `rubocop` v1.4.0

## 0.10.0 (2020-11-15)

### New Features

* Add `dependabot` config

### Changes

* Fix new warnings from `rubocop` v1.3.0 update
* `DocxParagraph#align` by default is
  symbol `:left`, not string `"left"`
* Store dev dependencies in `Gemfile.lock`
* Move all dev dependencies to `gemspec`
* Require ruby >= 2.5, since 2.4 EOLed

## 0.9.1 (2020-08-21)

### Fixes

* Fix `ParagraphMargins` side initialization in constructor

## 0.9.0 (2020-08-20)

### Changes

* Remove unused `HyperlinkForHover` class
* Replace `Chart.parse` method to `Chart#parse`
* Fix new warnings from `rubocop` v0.89.0 update
* Change `HslColor.rgb_to_hsl` to `Color#to_hsl`
* Remove deprecated `GridSpan#count_of_merged_cells`
  and `GridSpan#type`

## 0.8.1 (2020-08-04)

### Changes

* `RunProperties` can intialize `#font_name` in constructor

## 0.8.0 (2020-07-31)

### Fixes

* `Indents#to_s` result is single line

### Changes

* Move parsing `DocxDrawingProperties` inside class method
* Remove unused `Color#position` attr_accessor
* Remove unused `RunProperties#dirty` attr_accessor
* Changes from `rubocop` v0.88.0

## 0.7.2 (2020-07-10)

### Fixes

* Fix parsing `CellProperties#text_direction` value in pptx table

## 0.7.1 (2020-07-10)

### Fixes

* Fix parsing `RunProperties#baseline` value `superscript`

## 0.7.0 (2020-07-10)

### New Features

* Add basic support of parsing Pivot data (`PivotCache`,
  `PivotCacheDefinition`, `CacheSource`, `WorksheetSource`)
* Increase project test coverage
* New `OOXMLDocumentObject.encrypted_file?` param to ignore host-os
* Use GitHub Actions instead of Travis CI
* Add `markdownlint` support in GitHub Actions
* Add `rubocop` support in GitHub Actions
* Add support of `rubocop-rake`
* Add missing documentation
* Add GitHub action task to check 100% documented code
* Add `yard` gem as development dependency

### Fixes

* Do not raise waring if `FileReference#path` is correct url
* Fix comparing two child of `OOXMLDocumentObject` with different classes
* Fix `OoxmlSize#to_s` to output same result on all supported ruby-version

### Changes

* Drop support of ruby 2.3
* Remove `Picture` class alias to `DocxPicture`
* Simplify `TableStylePropertiesHelper` dynamic methods generation
* Remove `OldDocxShapeProperties#opacity` as unused
* Remove `CellProperties#anchor_center` as unused
* Remove `CellProperties#horizontal_overflow` as unused
* Remove `OldDocxPicture#style_number` as unused
* Remove `DocxShapeLineElement#type = quadratic_bezier` as unused
* Remove `OOXMLTextBox#properties` as unused
* Remove redundant comparing `Spacing` to `nil`
* Remove `DocxShapeProperties#text_box` as unused
* Remove parsing `Color#parse_color_model - scrgbClr` as unused
* Remove `DocxParagraphRun#shape - oval` as unused
* Remove `Shape#margin - right` as unused
* Remove `RunProperties#font_size_complex` as unused
* Remove `RunProperties#baseline - superscript` as unused
* Remove warning on `HeaderFooter#parse_type` on unknown type as unused
* Remove usage of `codecov` gem
* Remove codeclimate.com support
* Move `rubocop` dependencies in `gemspec` file
* Remove unused param from `DocxParagraphRun#parse_properties`
* Remove unused `Categories` alias for `SeriesText`
* Use `sh` command in `rake release_github_rubygems`

## 0.6.0 (2020-05-29)

### New Features

* Add support of `rubocop-performance`
* Add parsing `DocxShape#style` and it's child nodes
* Increase code coverage
* Simplify `Color#==`

### Changes

* Support of rubocop v0.84.0

### Fixes

* Fix coverage report on non-CI environments

## 0.5.1 (2020-03-24)

* Fix rake task for releasing gem

## 0.5.0 (2020-03-24)

### New features

* Do not use FileMagic on windows
* Parsing `FilterColumn#custom_filters`
* Parsing `Series#XYValues`
* Parsing `StructuredDocumentTag` as `TableCell#elements`
* Parsing `XlsxColumnProperties#hidden`
* Ability to parse shared strings in custom named file
* Parsing `CellProperties#borders_properties` for `tcBdr` tag
* Add detailed parsing of formulas in xlsx
* Redone parsing Bookmarks, since it part of
  `DocxParagraph` elements and order is important
* `BookmarkStart` and `BookmarkEnd` count as nonempty_run
* Add parsing of `CommentRangeStart` and `CommentRangeEnd`
* Add parsing of `Series#values`
* Add parsing of `DocumentStyle#default`
* Add parsing of `Worksheet#page_setup`
* Add parsing of `Worksheet#page_margins`
* Add parsing of `PageSetup#paper_size_name`
* Add parsing `TableProperties#fill`
* Add parsing `TablePart#TableColumns`
* Add parsing `Theme#fontScheme`
* Force strict parsing of XML to catch errors
* Add parsing `Chart#pivot_formats`
* Add parsing of `CommonDocumentStructure#ContentTypes`
* Add parsing `Chart#axis_ids`
* Add parsing `Chart#vary_colors`
* Add parsing `DocumentStructure#relationships`
* Add parsing `DocumentStructure#comments_document`
* Add parsing `ExcelComment#text`
* Add `Shade#to_s`
* Drop support of Ruby < 2.3
* `OoxmlSize#to_s` can output in different unit
* New rake task for release gem on github and rubygems

### Refactor

* Store `sdt` as `DocxParagraph#character_styles_array` element
* Deprecation warning for `DocxParagraph#frame_properties`
* Deprecation warning for `Point#text`
* Remove unused and probably not real `DocxParagraph#kinoku`
* Remove redundant `PresetColor`, `AbstractNumberingId`,
  `Start`, `VerticalMerge`, `Order`, `SeriesIndex`, `PointCount`,
  `Language`
* Refactor parsing `Chart`. Ability to parser multichart charts
* Refactor parsing `Color#parse_scheme_color`, no class method
* Refactor parsing `Color#parse_color_model`, no class method
* Refactor parsing `Color#parse_color`, no class method
* Refactor parsing `Borders#parse_from_style`, no class method
* Rename inlogical check for `DocxShapeLine#nil?` to DocxShapeLine#invisible?`
* Rename inlogical check for `ChartAxisTitle#nil?` to DocxShapeLine#visible?`
* Remove usage of `ThemeColors.list`
* Redone parsing shared string
* Change `Worksheet.parse` to instance method `Worksheet#parse`
* Change `XLSXWorkbook.parse` to instance method `XLSXWorkbook#parse`
* Remove class method `XLSXWorkbook.link_to_theme_xml`
* Remove class method `XLSXWorkbook.styles_node`
* Remove class `CellStyle`, replace with call of `style_sheet`
* Reorganize code for remove `Xf#calcualte_values` method
* Remove class methods during parsing excel comments
* Change `Presentation.default_font_size` to instance method `Presentation#default_font_size`
* Change `Presentation.default_font_typeface` to instance method `Presentation#default_font_typeface`
* Change parsing of table styles - move to `TableStyles` class
* Remove `Presentation.current_font_style`, use instance method instead
* Replace `Presentation.parse` and `Presentation#parse`
* Remove `Condition.parse_list`
* Redone parsing `XlsxCell`
* Redone parsing `CommenAuthors`
* Remove usage of deprecated `Coordinates#get_column_number`
* Use instance method `Relationships#parse_file` instead of class method `Relationships.parse_rels`
* Simplify parsing `DocxParagraphRun#font` for `asciiTheme`
* Redone parsing of `DocxStructure#comments`
* Change `DocxStructure#parse_default_style` to instance method
* Change `DocxStructure#parse_paragraph_style_xml` to instance method
* Redone parsing `Shape`
* Remove deprecated `GridSpan#type` and `GridSpan#count_of_merged_cells`
* Remove unused `ParagraphMargins#round`
* Remove unused `Spacing.default_spacing_canvas`
* Simplify `Spacing.parse_spacing_rule` into `LineSpacing` class
* `DocxShapeSize#rotation` now use OoxmlSize
* `OoxmlSize` support `one_60000th_degree` and `degree`
* `Columns#separator` is boolean

### Fixes

* Fix crash on empty coordinates list of chart
* Fix crash on docx with no `settings.xml`
* Fix crash on docx with no `theme.xml`
* Fix parsing hyperlinks with empty `id`
* Fix crashing on hover hyperlink in PPTX
* Fix incorrect values of Cell Number Format
* Fix `GridColumn#width` emu value
* Fix `DocumentStructure#with_data?` for docs with several empty paragraphs
* Fix `XLSXWorkbook#all_formula_values` for formulas without value
* Values of `DocxShapeSize#flip_horizontal` and
  `DocxShapeSize#flip_vertical` are Boolean
* Fix crash on gradient stop with unknown SchemeColor ([ooxml_parser#571](https://github.com/ONLYOFFICE/ooxml_parser/issues/571))
* Fix `Presentation.with_data?` for shape with preset geometry ([ooxml_parser#573](https://github.com/ONLYOFFICE/ooxml_parser/issues/573))
* Fix crash on parsing files without `styles.xml`
* Fix parsing audio `TimeNode`
* Fix class documentation issues
* Fix `Worksheet#with_data?` for columns without custom width

## 0.4.1 (2018-03-01)

### Fixes

* Fix calling `StructuredDocumentTag#parent`

## 0.4.0 (2018-02-21)

### New features

* Support `SDTContent#tables`

## 0.3.0 (2018-02-19)

### New features

* Add `to_hash` method to OOXMLDocumentObject. `to_json` also working
* Parsing `OleObject` in `GraphicFrame`
* Parsing `CommentsExtended`
* Parsing `DocxParagraph#paragraph_id`, `DocxParagraph#text_id`
* Parsing `ConnectionShape` in `Slide#elements`, `XlsxDrawing#shape`
* Parsing `TablePart#table_style_info`
* Parsing `Autofilter#filterColumn`
* Parsing `Worksheet#ole_objects`
* Parsing `AlternateContent#ole_object`
* Parsing `OOXMLShapeBodyProperties#number_columns`
* Parsing `OOXMLShapeBodyProperties#space_columns`
* Parsing `DocxParagraph#sdt`
* Parsing `DocumentStructure#sdt`
* Parsing `Presentation#relationships`
* Parsing `Extension#sparkline_groups`
* Parsing `SparklineGroup#type`, `SparklineGroup#line_weight`,
  `SparklineGroup` show points,
  `SparklineGroup` points colors, `SparklineGroup#display_empty_cells_as`,
  `SparklineGroup#display_hidden`, `SparklineGroup#display_x_axis`,
  `SparklineGroup#right_to_left`, SparklineGroup max and min axis type and value
* Parsing `PageProperties#title_page`
* Parsing `TableStyleProperties#table_row_properties`
* Parsing `ChartAxis#tick_label_position`
* Parsing `ChartAxis#scaling`
* Parsing `ConnectionShape` in `ShapesGrouping`
* Parsing `Tab#leader`
* Parsing Slide Notes
* Parsing `SDTProperties#alias`, `SDTProperties#lock`, `SDTProperties#tag`
* Parsing `Font#vertical_alignment`
* Parsing `Paragraph#hyperlink`
* Parsing `Hyperlink#runs`
* Parsing `ParagraphRun#t`
* Parsing `ParagraphRun#tab`

### Refactor

* Change interface of `XlsxDrawing#graphic_frame`
* Remove duplicate classes for `Tabs`
* Refactor `Table#autofilter`
* Simplification of `DocxParagraph#dup`
* Simplification of `DocxParagraphRun#dup`
* Simplification of `TableProperties#dup`
* Redone parsing of `Slide` data
* Remove `Color#parse_color_tag` in favor of `OoxmlColor#parse`
* Redone parsing `DocumentBackground#fill`
* Redone parsing and comparing `Color#style`

### Fixes

* Fix `uninitialized constant OoxmlParser::Inserted::DateTime` while parsing `Inserted#date`
* Fix detecting password protected files on old `FileMagic`
* `pct` OoxmlSize values actually is 1/50 of percent
* `DocumentStyle#paragraph_properties` is `ParagraphProperties`
* Fix crash if chart series has no values
* Fix crash if `Inserted#date` is incorrect
* Fix crash if `FileReference#resource_id` is empty
* Fix crash if `FileReference#path` is nil
* Do not hangup on parsing Coordinates like `Donut!A7:A7,Donut!A16:A16`
* Do not hangup on parsing Coordinates with `#`
* Fix crash on parsing Coordinates consist of `!`
* Fix crash on `Delimeter`
* Fix crash on parsing `Coordinates` with sheet name with `!`
* Fix checking `DocxParagraph#nonempty_runs` for runs with `Shape`

## 0.2.0 (2017-03-25)

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
* Remove usage of Linux `file` command in favor of
  `ruby-filemagick` gem. Better cross-platform support
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
* Parsing `TableStyleColumnBandSize`, `TableStyleRowBandSize`,
  `TableLayout`, table cell spacing for TableProperties
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
* Add parsing `DocxParagraphRun#object` and `Object#ole_object`
* New way to parse default RunProperties and
  ParagraphProperties. Old way is still there
* Add parsing `RunStyle`. Implement #140
* Add parsing `Chart#view3D`
* Add parsing math data in xlsx files
* Add parsing `PreSubSuperscript` class - `m:sPre` tag
* Add parsing `a:prstDash` for `DocxShapeLine`
* Add parsing `OOXMLShapeBodyProperties#vertical`
* Add parsing `CellStyle#apply_number_format`
* Add parsing `SheetView#show_gridlines`, `SheetView#show_row_column_headers`
* Do not crash, just show stderr if resource not found
* Add base support of `chartsheets`
* Add `parse_hex_string` for 3 digit colors
* Add parsing 'DocxParagraph#inserted' context
* Add parsing `bgPr` and `stretch`
* Add parsing `DisplayLabelsProperties` in `Series`
* Add parsing `FootnoteProperties`
* Add parsing `Settings#default_tab_stop`
* Add parsing `DocProperties`
* Add parsing `Table#description` and `Table#title`
* Add parsing `CommonNonVisualProperties#title` and `CommonNonVisualProperties#description`
* Add parsing `DocxPicture#non_visual_properties`
* Add parsing `GraphicFrame#non_visual_properties`
* Add parsing `X14Table`

### Fixes

* Fix parsing document style id - it can be string, not only digit
* Fix misplaced `dxa` and `emu` units of measurements and
  also fix calculation `dxa` unit
* Redone parsing of `nary` in formulas
* Fix parsing gradient color linear values
* Fix parsing Footnote and Endnote reference in runs
* Fix calculating position offset values, distance from
  text in different units of measurements
* Fix parsing table style in text box
* Fix problem with parsing absolute file path in Windows
* Parse `keep_next` in paragraph properties
* `TransformEffect`, `BordersProperties#size` in correct `OoxmlSize` unit
* `Indents`, `TableProperties#table_indent`, `ParagraphProperties#margin_left`,
  `ParagraphProperties#margin_right`, `ParagraphProperties#indent`,
  `DocxShapeLine#width`,
  `TextOutline#width`, `Outline#width`, `TableCellLine#width`,
  `XlsxDrawingPositionParameters` use `OoxmlSize`
* `TableMargins`, `TablePosition`, `FrameProperties`,
  `CellProperties#able_cell_width`, `TableRow#height` use `OoxmlSize`
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
* `DocxParagraph#nonempty_runs` recognized `footnote` and `endnote`
* Fix parsing emty `TransitionProperties`
* Fix memory hog on calculating diffs
* Fix values of `OOXMLFont`: was not float
* Fix parsing `Underline` style
* Fix parsing Strike `noStrike`
* Do not crash if indexed color have unknown index
* Fix converting to symbols for border styles
* Fix comparing docx's with shapes

### Refactor

* Redone parsing images - store in structure, instead of copying file to filesystem
* Move parsing Columns inside class
* Move parsing TableGrid inside class
* Move parsing DocxShapeLinePath inside class
* Totally redone parsing of Numbering
* Parsing DocxParagraphRun properties
* Move some slide methods to helper
* Rename `Size` class to `PageSize` and `RunSize` to `Size`
* Simplify code for parsing `CellProperties`, `PageProperties`,
  `DocxParagraph`, `Bookmark`, `TableLook`, `ChartAxisTitle`,
  `NonVisualProperties`, `ShapePlaceholder`, `RunProperties`,
  `TableProperties`, `FrameProperties`, `Background`, `Transition`,
  `TransitionProperties`, `XlsxRow`, `XlsxColumnProperties`, `CellStyle`,
  `DisplayLabelsProperties`, `ParagraphProperties`,
  `ChartAxis`, `ChartLegend`, `SizeRelativeHorizontal`,
  `SizeRelativeVertical`, `OOXMLShapeBodyProperties`, `Tile`,
  `NaryLimitLocation`, `MultilevelType`, `CellProperties`, `FileReference`, `OldDocxPicture`,
  `OldDocxShapeFill`, `ShapeGuide`, `ShapeAdjustValueList`,
  `PresetGeometry`, `DocxShapeSize`, `DocxShapeLinePath`,
  `DocxShapeLineElement`, `OOXMLCustomGeometry`, `DocxShapeProperties`,
  `LineEnd`, `CommonTiming`, `AnimationEffect`, `TargetElement`,
  `Behavior`, `SetTimeNode`, `Column`, `TableBorders`, `GridColumn`,
  `TableElement`, `XlsxAlignment`, `NaryProperties`, `NaryLimitLocation`,
  `NaryGrow`, `Nary`, `OOXMLFont`, `ExcelComment`, `SheetView`, `Pane`,
  `TablePart`
* Replace usage `Alignment.parse` on
  `OOXMLDocumentObject#value_to_symbol` method
* Refactor `Color.from_int16` to `Color#parse_hex_string` method
* Move `table_cell_spacing` to `TableRowProperties` and use OoxmlSize
* Simplify `XlsxRow` parsing.
* `FontStyle` class cleanup
* Simplify `Borders.parse_from_style`
* Simplify `Color.parse_color_tag`
* `Border` class cleanup
* Merge `PictureWidth` and `PictureHeight` in single `PictureDimension` class
* Redone parsing of `AlphaModFix` and blips

### Removal

* Remove unused method `PageSize.get_name_by_size`
* Remove unused class `TableStyleElement`
* Remove `OoxmlShift` - use `OoxmlCoordinates` instead
* Remove unused `Color#parse_int16_string`
* Remove parsing `TableProperties#right_to_left`, since it was totally wrong
* Remove class `ParagraphStyle`
* Remove useless attribute `OOXMLDocumentObject.namespace_perfix`
* Remove method `Color.parse_color_hash`

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
