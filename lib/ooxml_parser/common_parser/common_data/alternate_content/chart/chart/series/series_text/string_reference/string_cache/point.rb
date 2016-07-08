require_relative 'point/text_value'
module OoxmlParser
  # Class for parsing `c:pt` object
  class Point < OOXMLDocumentObject
    # @return [Integer] index of point
    attr_accessor :index
    # @return [TextValue] value of text
    attr_accessor :text

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

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'v'
          point.text = TextValue.parse(node_child)
        end
      end

      point
    end
  end
end
