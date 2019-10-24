# frozen_string_literal: true

require_relative 'shape_adjust_value_list/shape_guide'
module OoxmlParser
  # Class for describing List of Shape Adjust Values
  class ShapeAdjustValueList < OOXMLDocumentObject
    # @return [Array] list of shape guides
    attr_accessor :shape_guides_list

    def initialize(parent: nil)
      @shape_guides_list = []
      @parent = parent
    end

    # @return [Array, Column] accessor for relationship
    def [](key)
      @shape_guides_list[key]
    end

    # Parse ShapeAdjustValueList
    # @param [Nokogiri::XML:Node] node with List of Shape Adjust Values
    # @return [ShapeAdjustValueList] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'gd'
          @shape_guides_list << ShapeGuide.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
