require_relative 'shape_adjust_value_list/shape_guide'
module OoxmlParser
  # Class for describing List of Shape Adjust Values
  class ShapeAdjustValueList
    # @return [Array] list of shape guides
    attr_accessor :shape_guides_list

    def initialize
      @shape_guides_list = []
    end

    # @return [Array, Column] accessor for relationship
    def [](key)
      @shape_guides_list[key]
    end

    # Parse Relationships
    # @param [Nokogiri::XML:Node] node with List of Shape Adjust Values
    # @return [ShapeAdjustValueList] result of parsing
    def self.parse(node)
      list = ShapeAdjustValueList.new
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'gd'
          list.shape_guides_list << ShapeGuide.new(parent: list).parse(node_child)
        end
      end
      list
    end
  end
end
