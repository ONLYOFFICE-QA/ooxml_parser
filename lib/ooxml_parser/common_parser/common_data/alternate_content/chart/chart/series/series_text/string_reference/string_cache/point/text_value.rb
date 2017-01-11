module OoxmlParser
  # Class for parsing `c:v` object
  class TextValue < OOXMLDocumentObject
    # @return [String] text of value
    attr_accessor :value

    # Parse TextValue
    # @param [Nokogiri::XML:Node] node with TextValue
    # @return [TextValue] result of parsing
    def parse(node)
      @value = node.text
      self
    end
  end
end
