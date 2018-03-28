require_relative 'string_cache/point'
module OoxmlParser
  # Class for parsing `c:tx`, `c:numCache` object
  class NumberStringCache < OOXMLDocumentObject
    # @return [String] Format Code
    attr_reader :format_code
    # @return [StringReference] String reference of series
    attr_reader :point_count
    # @return [Array, Point] array of points
    attr_reader :points

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
        when 'formatCode'
          @format_code = node_child.text
        when 'ptCount'
          @point_count = ValuedChild.new(:integer, parent: self).parse(node_child)
        when 'pt'
          @points << Point.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
