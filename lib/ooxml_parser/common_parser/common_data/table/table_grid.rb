# frozen_string_literal: true

require_relative 'table_grid/grid_column'
module OoxmlParser
  # Class for parsing `w:tblGrid` object
  class TableGrid < OOXMLDocumentObject
    # @return [Array, GridColumn] array of columns
    attr_accessor :columns

    def initialize(parent: nil)
      @columns = []
      @parent = parent
    end

    def ==(other)
      @columns.each_with_index do |cur_column, index|
        return false unless cur_column == other.columns[index]
      end
      true
    end

    # Parse TableGrid
    # @param [Nokogiri::XML:Node] node with TableGrid
    # @return [TableGrid] result of parsing
    def parse(node)
      node.xpath('*').each do |grid_child|
        case grid_child.name
        when 'gridCol'
          @columns << GridColumn.new(parent: self).parse(grid_child)
        end
      end
      self
    end
  end
end
