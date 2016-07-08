require_relative 'string_cache/point_count'
require_relative 'string_cache/point'
module OoxmlParser
  # Class for parsing `c:tx` object
  class StringCache < OOXMLDocumentObject
    # @return [StringReference] String reference of series
    attr_accessor :point_count
    # @return [Array, Point] array of points
    attr_accessor :points

    def initialize
      @points = []
    end

    # Parse Order
    # @param [Nokogiri::XML:Node] node with Order
    # @return [Order] result of parsing
    def self.parse(node)
      cache = StringCache.new
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'ptCount'
          cache.point_count = PointCount.parse(node_child)
        when 'pt'
          cache.points << Point.parse(node_child)
        end
      end
      cache
    end
  end
end
