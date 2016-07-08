require_relative 'string_cache/point_count'
module OoxmlParser
  # Class for parsing `c:tx` object
  class StringCache < OOXMLDocumentObject
    # @return [StringReference] String reference of series
    attr_accessor :point_count

    # Parse Order
    # @param [Nokogiri::XML:Node] node with Order
    # @return [Order] result of parsing
    def self.parse(node)
      cache = StringCache.new
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'ptCount'
          cache.point_count = PointCount.parse(node_child)
        end
      end
      cache
    end
  end
end
