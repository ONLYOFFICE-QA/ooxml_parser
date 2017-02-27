require_relative 'cell_style/foreground_color'
require_relative 'cell_style/alignment'
require_relative 'cell_style/ooxml_font'
module OoxmlParser
  # Class for parsing cell style
  class CellStyle < OOXMLDocumentObject
    ALL_FORMAT_VALUE = %w|0 0.00 #,##0 #,##0.00 $#,##0_);($#,##0) $#,##0_);[Red]($#,##0) $#,##0.00_);($#,##0.00)
                          $#,##0.00_);[Red]($#,##0.00)
                          0.00% 0.00%
                          0.00E+00
                          #\ ?/?
                          #\ ??/??
                          m/d/yyyy
                          d-mmm-yy
                          d-mmm
                          mmm-yy
                          h:mm AM/PM
                          h:mm:ss AM/PM
                          h:mm
                          h:mm:ss
                          m/d/yyyy h:mm
                          0
                          0
                          0
                          0
                          0
                          0
                          0
                          0
                          0
                          0
                          0
                          0
                          0
                          0
                          #,##0_);(#,##0)
                          #,##0_);[Red](#,##0)
                          #,##0.00_);(#,##0.00)
                          #,##0.00_);[Red](#,##0.00)
                          0
                          0
                          0
                          0
                          mm:ss
                          General
                          mm:ss.0
                          ##0.0E+0
                          @|.freeze

    attr_accessor :font, :borders, :fill_color, :numerical_format, :alignment
    # @return [True, False] check if style should add QuotePrefix (' symbol) to start of the string
    attr_accessor :quote_prefix
    # @return [True, False] is font applied
    attr_accessor :apply_font
    # @return [True, False] is border applied
    attr_accessor :apply_border
    # @return [True, False] is fill applied
    attr_accessor :apply_fill
    # @return [True, False] is number format applied
    attr_accessor :apply_number_format
    # @return [True, False] is alignment applied
    attr_accessor :apply_alignment
    # @return [Integer] id of font
    attr_accessor :font_id
    # @return [Integer] id of border
    attr_accessor :border_id
    # @return [Integer] id of fill
    attr_accessor :fill_id
    # @return [Integer] id of number format
    attr_accessor :number_format_id

    def initialize(parent: nil)
      @numerical_format = 'General'
      @alignment = XlsxAlignment.new
      @parent = parent
    end

    # Parse CellStyle object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [CellStyle] result of parsing
    def parse(node)
      style_node = XLSXWorkbook.styles_node.xpath('//xmlns:cellXfs/xmlns:xf')[node.to_i]
      style_node.attributes.each do |key, value|
        case key
        when 'applyFont'
          @apply_font = attribute_enabled?(value)
        when 'applyBorder'
          @apply_border = attribute_enabled?(value)
        when 'applyFill'
          @apply_fill = attribute_enabled?(value)
        when 'applyNumberFormat'
          @apply_number_format = attribute_enabled?(value)
        when 'applyAlignment'
          @apply_alignment = attribute_enabled?(value)
        when 'fontId'
          @font_id = value.value.to_i
        when 'borderId'
          @border_id = value.value.to_i
        when 'fillId'
          @fill_id = value.value.to_i
        when 'numFmtId'
          @number_format_id = value.value.to_i
        when 'quotePrefix'
          @quote_prefix = attribute_enabled?(value)
        end
      end
      style_node.xpath('*').each do |node_child|
        case node_child.name
        when 'alignment'
          @alignment.parse(node_child) if @apply_alignment
        end
      end
      calculate_values
      self
    end

    def calculate_values
      @font = if @apply_font
                OOXMLFont.new(parent: self).parse(0)
              else
                OOXMLFont.new(parent: self).parse(@font_id)
              end
      @borders = Borders.parse_from_style(@border_id) if @apply_border
      @fill_color = ForegroundColor.new(parent: self).parse(@fill_id) if @apply_fill
      return unless @apply_number_format
      XLSXWorkbook.styles_node.xpath('//xmlns:numFmt').each do |numeric_format|
        if @number_format_id == numeric_format.attribute('numFmtId').value.to_i
          @numerical_format = numeric_format.attribute('formatCode').value
        elsif CellStyle::ALL_FORMAT_VALUE[@number_format_id - 1]
          @numerical_format = CellStyle::ALL_FORMAT_VALUE[@number_format_id - 1]
        end
      end
      @numerical_format = CellStyle::ALL_FORMAT_VALUE[@number_format_id - 1] if CellStyle::ALL_FORMAT_VALUE[@number_format_id - 1]
    end
  end
end
