module OoxmlParser
  # Class for parsing `c:ptCount` object
  class PointCount < OOXMLDocumentObject
    # @return [Integer] value of point count
    attr_accessor :value

    # Parse PointCount
    # @param [Nokogiri::XML:Node] node with PointCount
    # @return [PointCount] result of parsing
    def self.parse(node)
      point_count = PointCount.new
      node.attributes.each do |key, value|
        case key
        when 'val'
          point_count.value = value.value.to_f
        end
      end
      point_count
    end
  end
end
