module OoxmlParser
  # Properties of XLSX column
  class XlsxColumnProperties < OOXMLDocumentObject
    attr_accessor :from, :to, :width, :style

    # Parse XlsxColumnProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [XlsxColumnProperties] result of parsing
    def parse(col_node)
      @from = col_node.attribute('min').value.to_i
      @to = col_node.attribute('max').value.to_i
      @style = CellStyle.parse(col_node.attribute('style').value) unless col_node.attribute('style').nil?
      @width = col_node.attribute('width').value.to_f - 0.7109375 if option_enabled?(col_node, 'customWidth')
      self
    end

    def self.parse_list(columns_width_node)
      columns = []
      columns_width_node.xpath('xmlns:col').each do |col_node|
        col = XlsxColumnProperties.new(parent: columns).parse(col_node)
        columns << col
      end
      columns
    end
  end
end
