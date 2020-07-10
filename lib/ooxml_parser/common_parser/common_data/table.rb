# frozen_string_literal: true

require_relative 'table/row/row'
require_relative 'table/table_properties'
require_relative 'table/table_grid'
require_relative 'table/margins/table_margins'
require_relative 'table/margins/paragraph_margins'
module OoxmlParser
  # Class for Table data
  class Table < OOXMLDocumentObject
    attr_accessor :grid, :rows, :properties, :number

    def initialize(rows = [], parent: nil)
      @rows = rows
      @parent = parent
    end

    alias table_properties properties

    # @return [String] result of convert of object to string
    def to_s
      "Rows: #{@rows.join(',')}"
    end

    # @return [String] inspect of object for debug means
    def inspect
      to_s
    end

    # Parse Table object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Table] result of parsing
    def parse(node,
              number = 0,
              default_table_properties = TableProperties.new)
      table_properties = default_table_properties.dup
      table_properties.jc = :left
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'tblGrid'
          @grid = TableGrid.new(parent: self).parse(node_child)
        when 'tr'
          @rows << TableRow.new(parent: self).parse(node_child)
        when 'tblPr'
          @properties = TableProperties.new(parent: self).parse(node_child)
        end
      end
      @number = number
      self
    end
  end
end
