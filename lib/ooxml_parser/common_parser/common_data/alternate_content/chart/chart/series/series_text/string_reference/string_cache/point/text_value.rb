module OoxmlParser
  # Class for parsing `c:v` object
  class TextValue < OOXMLDocumentObject
    # @return [String] text of value
    attr_accessor :value

    # Parse TextValue
    # @param [Nokogiri::XML:Node] node with TextValue
    # @return [TextValue] result of parsing
    def self.parse(node)
      text_value = TextValue.new
      text_value.value = node.text
      text_value
    end
  end
end
