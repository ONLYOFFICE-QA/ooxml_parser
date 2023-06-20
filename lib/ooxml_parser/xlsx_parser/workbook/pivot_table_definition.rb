# frozen_string_literal: true

require_relative 'pivot_table_definition/column_row_items'
require_relative 'pivot_table_definition/location'
require_relative 'pivot_table_definition/pivot_fields'
require_relative 'pivot_table_definition/pivot_table_style_info'
require_relative 'pivot_table_definition/data_fields'
require_relative 'pivot_table_definition/page_fields'

module OoxmlParser
  # Class for parsing <PivotTableDefinition> tag
  class PivotTableDefinition < OOXMLDocumentObject
    # @return [String] name of table
    attr_reader :name
    # @return [Integer] id of cache
    attr_reader :cache_id
    # @return [True, False] should number formats be applied
    attr_reader :apply_number_formats
    # @return [True, False] should border formats be applied
    attr_reader :apply_border_formats
    # @return [True, False] should font formats be applied
    attr_reader :apply_font_formats
    # @return [True, False] should pattern formats be applied
    attr_reader :apply_pattern_formats
    # @return [True, False] should alignment formats be applied
    attr_reader :apply_alignment_formats
    # @return [True, False] should width height formats be applied
    attr_reader :apply_width_height_formats
    # @return [True, False] should auto formatting be used
    attr_reader :use_auto_formatting
    # @return [True, False] should item print titles
    attr_reader :item_print_titles
    # @return [String] data caption
    attr_reader :data_caption
    # @return [Integer] creation version
    attr_reader :created_version
    # @return [Integer] indent
    attr_reader :indent
    # @return [True, False] outline
    attr_reader :outline
    # @return [True, False] outline data
    attr_reader :outline_data
    # @return [True, False] is there multiple fields filters
    attr_reader :multiple_field_filters
    # @return [Location] location data
    attr_reader :location
    # @return [PivotFields] pivot fields
    attr_reader :pivot_fields
    # @return [ColumnRowItems] column items
    attr_reader :column_items
    # @return [ColumnRowItems] row items
    attr_reader :row_items
    # @return [PivotTableStyleInfo] style info
    attr_reader :style_info
    # @return [DataFields] data fields
    attr_reader :data_fields
    # @return [PageFields] page fields
    attr_reader :page_fields

    # Parse PivotTableDefinition object
    # @param [String] file path
    # @return [PivotTableDefinition] result of parsing
    def parse(file)
      doc = parse_xml("#{root_object.unpacked_folder}/#{file}")
      node = doc.xpath('//xmlns:pivotTableDefinition').first
      node.attributes.each do |key, value|
        case key
        when 'name'
          @name = value.value.to_s
        when 'cacheId'
          @cache_id = value.value.to_i
        when 'applyNumberFormats'
          @apply_number_formats = attribute_enabled?(value)
        when 'applyBorderFormats'
          @apply_border_formats = attribute_enabled?(value)
        when 'applyFontFormats'
          @apply_font_formats = attribute_enabled?(value)
        when 'applyPatternFormats'
          @apply_pattern_formats = attribute_enabled?(value)
        when 'applyAlignmentFormats'
          @apply_alignment_formats = attribute_enabled?(value)
        when 'applyWidthHeightFormats'
          @apply_width_height_formats = attribute_enabled?(value)
        when 'useAutoFormatting'
          @use_auto_formatting = attribute_enabled?(value)
        when 'itemPrintTitles'
          @item_print_titles = attribute_enabled?(value)
        when 'dataCaption'
          @data_caption = value.value.to_s
        when 'createdVersion'
          @created_version = value.value.to_i
        when 'indent'
          @indent = value.value.to_i
        when 'outline'
          @outline = attribute_enabled?(value)
        when 'outlineData'
          @outline_data = attribute_enabled?(value)
        when 'multipleFieldFilters'
          @multiple_field_filters = attribute_enabled?(value)
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'location'
          @location = Location.new(parent: self).parse(node_child)
        when 'pivotFields'
          @pivot_fields = PivotFields.new(parent: self).parse(node_child)
        when 'rowItems'
          @row_items = ColumnRowItems.new(parent: self).parse(node_child)
        when 'colItems'
          @column_items = ColumnRowItems.new(parent: self).parse(node_child)
        when 'pivotTableStyleInfo'
          @style_info = PivotTableStyleInfo.new(parent: self).parse(node_child)
        when 'dataFields'
          @data_fields = DataFields.new(parent: self).parse(node_child)
        when 'pageFields'
          @page_fields = PageFields.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
