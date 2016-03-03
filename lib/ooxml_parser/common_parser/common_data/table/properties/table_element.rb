require_relative 'whole_table/table_cell_style'
# Describe single table element
module OoxmlParser
  class TableElement
    attr_accessor :cell_style

    alias cell_properties cell_style

    def self.parse(whole_table_node)
      table_element = TableElement.new
      whole_table_node.xpath('*').each do |whole_table_node_child|
        case whole_table_node_child.name
        when 'tcStyle', 'tcPr'
          table_element.cell_style = CellProperties.parse(whole_table_node_child)
        end
      end
      table_element
    end
  end
end
