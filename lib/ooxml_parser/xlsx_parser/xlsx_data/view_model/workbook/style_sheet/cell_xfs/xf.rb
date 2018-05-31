require_relative 'cell_style/alignment'
module OoxmlParser
  # Class for parsing `xf` object
  class Xf < OOXMLDocumentObject
    ALL_FORMAT_VALUE = ['General',
                        '0',
                        '0.00',
                        '#,##0',
                        '#,##0.00',
                        '$#,##0_);($#,##0)',
                        '$#,##0_);[Red]($#,##0)',
                        '$#,##0.00_);($#,##0.00)',
                        '$#,##0.00_);[Red]($#,##0.00)',
                        '0%', '0.00%',
                        '0.00E+00',
                        '# ?/?',
                        '# ??/??',
                        'm/d/yyyy',
                        'd-mmm-yy',
                        'd-mmm',
                        'mmm-yy',
                        'h:mm AM/PM',
                        'h:mm:ss AM/PM',
                        'h:mm',
                        'h:mm:ss',
                        'm/d/yyyy h:mm',
                        nil,
                        nil,
                        nil,
                        nil,
                        nil,
                        nil,
                        nil,
                        nil,
                        nil,
                        nil,
                        nil,
                        nil,
                        nil,
                        nil,
                        '#,##0_);(#,##0)',
                        '#,##0_);[Red](#,##0)',
                        '#,##0.00_);(#,##0.00)',
                        '#,##0.00_);[Red](#,##0.00)',
                        nil,
                        nil,
                        nil,
                        nil,
                        'mm:ss',
                        '[h]:mm:ss',
                        'mm:ss.0',
                        '##0.0E+0',
                        '@'].freeze

    attr_reader :font, :borders, :fill_color, :numerical_format, :alignment
    # @return [True, False] check if style should add QuotePrefix (' symbol) to start of the string
    attr_reader :quote_prefix
    # @return [True, False] is font applied
    attr_reader :apply_font
    # @return [True, False] is border applied
    attr_reader :apply_border
    # @return [True, False] is fill applied
    attr_reader :apply_fill
    # @return [True, False] is number format applied
    attr_reader :apply_number_format
    # @return [True, False] is alignment applied
    attr_reader :apply_alignment
    # @return [Integer] id of font
    attr_reader :font_id
    # @return [Integer] id of border
    attr_reader :border_id
    # @return [Integer] id of fill
    attr_reader :fill_id
    # @return [Integer] id of number format
    attr_reader :number_format_id

    def initialize(parent: nil)
      @numerical_format = 'General'
      @alignment = XlsxAlignment.new
      @parent = parent
    end

    # Parse Xf object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Xf] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
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
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'alignment'
          @alignment.parse(node_child) if @apply_alignment
        end
      end
      self
    end

    def calculate_values
      @font = root_object.style_sheet.fonts[@font_id]
      @borders = root_object.style_sheet.borders.borders_array[@border_id] if @apply_border
      @fill_color = root_object.style_sheet.fills[@fill_id] if @apply_fill
      return self unless @apply_number_format
      format = root_object.style_sheet.number_formats.format_by_id(@number_format_id)
      @numerical_format = if format
                            format.format_code
                          else
                            ALL_FORMAT_VALUE[@number_format_id]
                          end
      self
    end
  end
end
