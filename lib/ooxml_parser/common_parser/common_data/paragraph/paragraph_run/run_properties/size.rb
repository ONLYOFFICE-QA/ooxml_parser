module OoxmlParser
  # Class for parsing `w:sz` object
  class Size < OOXMLDocumentObject
    # @return [Integer] value of size
    attr_accessor :value

    # Parse Size
    # @param [Nokogiri::XML:Node] node with Size
    # @return [Size] result of parsing
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
