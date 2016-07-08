module OoxmlParser
  # Class for parsing `c:pt` object
  class Point < OOXMLDocumentObject
    # @return [Integer] index of point
    attr_accessor :index

    # Parse PointCount
    # @param [Nokogiri::XML:Node] node with PointCount
    # @return [PointCount] result of parsing
    def self.parse(node)
      point = Point.new
      node.attributes.each do |key, value|
        case key
        when 'idx'
          point.index = value.value.to_f
        end
      end
      point
    end
  end
end
