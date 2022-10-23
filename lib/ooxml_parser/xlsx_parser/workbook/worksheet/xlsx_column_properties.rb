# frozen_string_literal: true

module OoxmlParser
  # Properties of XLSX column
  class XlsxColumnProperties < OOXMLDocumentObject
    attr_accessor :style
    # @return [True, False] is width custom
    attr_accessor :custom_width
    # @return [True, False] Flag indicating if the
    # specified column(s) is set to 'best fit'
    attr_accessor :best_fit
    # @return [True, False] Flag indicating if the
    # specified column(s) is hidden
    attr_reader :hidden
    # @return [Integer] First column affected by this 'column info' record.
    attr_reader :min
    # @return [Integer] Last column affected by this 'column info' record.
    attr_reader :max
    # @return [Float] width in pixel, as stored in xml structure
    attr_reader :width_raw
    # @return [Float] width, in readable format
    attr_reader :width

    # Parse XlsxColumnProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [XlsxColumnProperties] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'min'
          @min = value.value.to_i
        when 'max'
          @max = value.value.to_i
        when 'style'
          @style = root_object.style_sheet.cell_xfs.xf_array[value.value.to_i]
        when 'width'
          @width_raw = value.value.to_f
          @width = calculate_width(@width_raw)
        when 'customWidth'
          @custom_width = option_enabled?(node, 'customWidth')
        when 'bestFit'
          @best_fit = boolean_attribute_value(value)
        when 'hidden'
          @hidden = boolean_attribute_value(value)
        end
      end
      self
    end

    alias from min
    alias to max

    extend Gem::Deprecate
    deprecate :from, 'min', 2099, 1
    deprecate :to, 'max', 2099, 1

    # Parse list of XlsxColumnProperties
    # @param columns_width_node [Nokogiri::XML:Element] node to parse
    # @param parent [OOXMLDocumentObject] parent of result objects
    # @return [Array<XlsxColumnProperties>] list of XlsxColumnProperties
    def self.parse_list(columns_width_node, parent: nil)
      columns = []
      columns_width_node.xpath('xmlns:col').each do |col_node|
        col = XlsxColumnProperties.new(parent: parent).parse(col_node)
        columns << col
      end
      columns
    end

    private

    # TODO: Currently width calculation use some magick number from old time ago
    #   Actual formula is way more complicated and require data about
    #   font max width for single character.
    #   Read more in ECMA-376, §18.3.1.13
    # @param [Float] value raw value for width
    # @return [Float] width value
    def calculate_width(value)
      value - 0.7109375
    end
  end
end
