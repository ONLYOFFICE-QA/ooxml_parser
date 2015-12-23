require_relative 'xlsx_row/xlsx_cell'
# Single Row of XLSX
module OoxmlParser
  class XlsxRow < OOXMLDocumentObject
    attr_accessor :cells, :height, :style, :hidden

    def initialize(cells = [])
      @cells = cells
    end

    def self.parse(row_node)
      row = XlsxRow.new
      row.height = row_node.attribute('ht').value if OOXMLDocumentObject.option_enabled?(row_node, 'customHeight')
      row.hidden = OOXMLDocumentObject.option_enabled?(row_node, 'hidden')
      row_node.xpath('xmlns:c').each { |cell_node| row.cells[Coordinates.parse_coordinates_from_string(cell_node.attribute('r').value.to_s).get_column_number.to_i - 1] = XlsxCell.parse(cell_node) }
      row
    end
  end
end
