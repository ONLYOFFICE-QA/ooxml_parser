module OoxmlParser
  # Class for parsing `w:position` object
  class Position < OOXMLDocumentObject
    # @return [String] value of position
    attr_accessor :value

    # Parse Position
    # @param [Nokogiri::XML:Node] node with Position
    # @return [Position] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = OoxmlSize.new(value.value.to_f, :half_point)
        end
      end
      self
    end
  end
end
