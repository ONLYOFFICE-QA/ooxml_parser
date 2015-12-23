require_relative 'line_join'
# Table Cell Style in XLSX
module OoxmlParser
  class TableCellStyle
    attr_accessor :borders, :fill

    def self.parse(cell_style_node)
      cell_style = TableCellStyle.new
      cell_style_node.xpath('*').each do |cell_node_child|
        case cell_node_child.name
        when 'tcBdr'
          cell_style.borders = Borders.parse(cell_node_child)
        when 'fill'
          cell_style.fill = PresentationFill.parse(cell_node_child)
        end
      end
      cell_style
    end
  end
end
