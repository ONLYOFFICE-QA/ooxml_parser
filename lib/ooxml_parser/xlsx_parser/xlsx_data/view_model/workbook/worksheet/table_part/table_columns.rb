# frozen_string_literal: true

require_relative 'table_columns/table_column'
module OoxmlParser
  # Class for `tableColumns` tag data
  class TableColumns < OOXMLDocumentObject
    # @return [Integer] count of columns
    attr_reader :count
    # @return [Array<TableColumn>] array of columns
    attr_reader :table_column_array

    def initialize(parent: nil)
      @table_column_array = []
      @parent = parent
    end

    # @return [Array, TableColumn] accessor
    def [](key)
      @table_column_array[key]
    end

    # Parse Table Columns data
    # @param [Nokogiri::XML:Element] node with TableColumns data
    # @return [ExtensionList] value of TableColumns data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'count'
          @count = value.value.to_i
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'tableColumn'
          @table_column_array << TableColumn.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
