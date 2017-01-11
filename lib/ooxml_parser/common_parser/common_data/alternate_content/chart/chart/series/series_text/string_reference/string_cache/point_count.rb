module OoxmlParser
  # Class for parsing `c:ptCount` object
  class PointCount < OOXMLDocumentObject
    # @return [Integer] value of point count
    attr_accessor :value

    # Parse PointCount
    # @param [Nokogiri::XML:Node] node with PointCount
    # @return [PointCount] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = value.value.to_f
        end
      end
      self
    end
  end
end
