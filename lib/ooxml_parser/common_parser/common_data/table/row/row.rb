require_relative 'cell/cell'
module OoxmlParser
  class TableRow < OOXMLDocumentObject
    attr_accessor :height, :cells, :cells_spacing

    def initialize(cells = [])
      @cells = cells
    end

    alias table_row_height height

    def self.parse(row_node)
      row = TableRow.new
      row.height = (row_node.attribute('h').value.to_f / 360_000.0).round(2) unless row_node.attribute('h').nil?
      row_node.xpath('*').each do |row_node_child|
        row_node_child.xpath('*').each do |row_properties|
          case row_properties.name
          when 'tblCellSpacing'
            row.cells_spacing = (row_properties.attribute('w').value.to_f / 283.3).round(2)
          end
        end
      end
      Presentation.current_font_style = FontStyle.new(true) # TODO: Add correct parsing of TableStyle.xml file and use it
      row_node.xpath("#{OOXMLDocumentObject.namespace_prefix}:tc").each { |cell_node| row.cells << TableCell.parse(cell_node) }
      Presentation.current_font_style = FontStyle.new
      row
    end
  end
end
