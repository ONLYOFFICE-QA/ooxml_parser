# Table Style in XLSX
module OoxmlParser
  class TableStyleElement
    attr_accessor :cell_properties

    def initialize
      @cell_properties = CellProperties.new
    end

    def self.parse(style_node)
      element_style = TableStyleElement.new
      style_node.xpath('*').each do |style_node_child|
        case style_node_child.name
        when 'tcPr'
          element_style.cell_properties = CellProperties.parse(style_node_child)
        end
      end
      element_style
    end
  end
end
