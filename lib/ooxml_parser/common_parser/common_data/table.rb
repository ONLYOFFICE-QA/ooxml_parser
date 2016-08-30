require_relative 'table/row/row'
require_relative 'table/table_properties'
require_relative 'table/table_grid'
require_relative 'table/margins/table_margins'
require_relative 'table/margins/paragraph_margins'
module OoxmlParser
  class Table < OOXMLDocumentObject
    attr_accessor :grid, :rows, :properties, :number

    def initialize(rows = [])
      @rows = rows
    end

    alias table_properties properties

    def self.parse(table_node,
                   number = 0,
                   default_table_properties = TableProperties.new,
                   parent: nil)
      table_properties = default_table_properties.copy
      table_properties.jc = :left
      table = Table.new
      table.parent = parent
      table_node.xpath('*').each do |table_node_child|
        case table_node_child.name
        when 'tblGrid'
          table.grid = TableGrid.parse(table_node_child)
        when 'tr'
          table.rows << TableRow.parse(table_node_child, parent: table)
        when 'tblPr'
          table.properties = TableProperties.new(parent: table).parse(table_node_child)
        end
      end
      table.number = number
      table
    end
  end
end
