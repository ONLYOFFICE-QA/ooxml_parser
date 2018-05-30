module OoxmlParser
  # Properties of XLSX column
  class XlsxColumnProperties < OOXMLDocumentObject
    attr_accessor :from, :to, :width, :style
    # @return [True, False] is width custom
    attr_accessor :custom_width
    # @return [True, False] Flag indicating if the
    # specified column(s) is set to 'best fit'
    attr_accessor :best_fit
    # @return [True, False] Flag indicating if the
    # specified column(s) is hidden
    attr_reader :hidden

    # Parse XlsxColumnProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [XlsxColumnProperties] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'min'
          @from = value.value.to_i
        when 'max'
          @to = value.value.to_i
        when 'style'
          @style = root_object.style_sheet.cell_xfs.xf_array[value.value.to_i].calculate_values
        when 'width'
          @width = value.value.to_f - 0.7109375
        when 'customWidth'
          @custom_width = option_enabled?(node, 'customWidth')
        when 'bestFit'
          @best_fit = attribute_enabled?(value)
        when 'hidden'
          @hidden = attribute_enabled?(value)
        end
      end
      self
    end

    def self.parse_list(columns_width_node, parent: nil)
      columns = []
      columns_width_node.xpath('xmlns:col').each do |col_node|
        col = XlsxColumnProperties.new(parent: parent).parse(col_node)
        columns << col
      end
      columns
    end
  end
end
