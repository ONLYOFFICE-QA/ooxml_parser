# Properties of XLSX column
module OoxmlParser
  class XlsxColumnProperties < OOXMLDocumentObject
    attr_accessor :from, :to, :width, :style

    def self.parse(col_node)
      col = XlsxColumnProperties.new
      col.from = col_node.attribute('min').value.to_i
      col.to = col_node.attribute('max').value.to_i
      col.style = CellStyle.parse(col_node.attribute('style').value) unless col_node.attribute('style').nil?
      col.width = col_node.attribute('width').value.to_f - 0.7109375 if OOXMLDocumentObject.option_enabled?(col_node, 'customWidth')
      col
    end

    def self.parse_list(columns_width_node)
      columns = []
      columns_width_node.xpath('xmlns:col').each do |col_node|
        col = XlsxColumnProperties.parse(col_node)
        columns << col
      end
      columns
    end
  end
end
