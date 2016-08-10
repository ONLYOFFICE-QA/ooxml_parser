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
                   default_paragraph = DocxParagraph.new,
                   default_run = DocxParagraphRun.new,
                   parent: nil)
      table_properties = default_table_properties.copy
      table_properties.jc = :left
      table_paragraph = default_paragraph.copy
      table_character = default_run.copy
      table = Table.new
      table.parent = parent
      table_node.xpath('*').each do |table_node_child|
        case table_node_child.name
        when 'tblGrid'
          table.grid = TableGrid.parse(table_node_child)
        when 'tr'
          table.rows << TableRow.parse(table_node_child, parent: table)
        when 'tblPr'
          table.properties = TableProperties.parse(table_node_child, parent: table)
        when ''
          DocxParagraph.parse_paragraph_style(table_node_child, table_paragraph, table_character)
        end
      end
      table.number = number
      table
    end
  end
end
