require_relative 'string_cache/point_count'
require_relative 'string_cache/point'
module OoxmlParser
  # Class for parsing `c:tx` object
  class StringCache < OOXMLDocumentObject
    # @return [StringReference] String reference of series
    attr_accessor :point_count
    # @return [Array, Point] array of points
    attr_accessor :points

    def initialize(parent: nil)
      @points = []
      @parent = parent
    end

    # Parse Order
    # @param [Nokogiri::XML:Node] node with Order
    # @return [Order] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'ptCount'
          @point_count = PointCount.new(parent: self).parse(node_child)
        when 'pt'
          @points << Point.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
