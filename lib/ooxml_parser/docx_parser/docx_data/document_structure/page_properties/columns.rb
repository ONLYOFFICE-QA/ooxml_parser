# frozen_string_literal: true

require_relative 'columns/column'
module OoxmlParser
  class Columns < OOXMLDocumentObject
    attr_accessor :separator
    attr_accessor :count
    attr_accessor :equal_width
    alias equal_width? equal_width
    attr_accessor :column_array
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
          @separator = value.value
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
