# frozen_string_literal: true

require_relative 'point/text_value'
module OoxmlParser
  # Class for parsing `c:pt` object
  class Point < OOXMLDocumentObject
    # @return [Integer] index of point
    attr_accessor :index
    # @return [TextValue] value of text
    attr_reader :value

    alias text value

    extend Gem::Deprecate
    deprecate :text, 'OoxmlParser::Point#value', 2020, 1

    # Parse PointCount
    # @param [Nokogiri::XML:Node] node with PointCount
    # @return [PointCount] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'idx'
          @index = value.value.to_f
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'v'
          @value = TextValue.new(parent: self).parse(node_child)
        end
      end

      self
    end
  end
end
