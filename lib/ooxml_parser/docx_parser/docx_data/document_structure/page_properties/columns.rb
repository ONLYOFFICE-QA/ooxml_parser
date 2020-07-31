# frozen_string_literal: true

require_relative 'columns/column'
module OoxmlParser
  # Class for data of Columns
  class Columns < OOXMLDocumentObject
    # @return [Integer] count of columns
    attr_accessor :count
    # @return [True, False] is columns are equal width
    attr_accessor :equal_width
    alias equal_width? equal_width
    # @return [Array<Column>] list of colujmns
    attr_accessor :column_array
    # @return [Boolean] Draw Line Between Columns
    attr_reader :separator
    # @return [OoxmlSize] space between columns
    attr_accessor :space

    def initialize(columns_count = 0)
      @count = columns_count
      @column_array = []
    end

    # @return [Array, Column] accessor
    def [](key)
      @column_array[key]
    end

    # Parse Columns data
    # @param [Nokogiri::XML:Element] node with Columns data
    # @return [Columns] value of Columns data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'num'
          @count = value.value.to_i
        when 'sep'
          @separator = option_enabled?(node, 'sep')
        when 'equalWidth'
          @equal_width = option_enabled?(node, 'equalWidth')
        when 'space'
          @space = OoxmlSize.new(value.value.to_f)
        end
      end

      node.xpath('*').each do |column_node|
        case column_node.name
        when 'col'
          @column_array << Column.new(parent: self).parse(column_node)
        end
      end
      self
    end
  end
end
