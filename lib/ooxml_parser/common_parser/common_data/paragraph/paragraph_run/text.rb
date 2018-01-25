module OoxmlParser
  # Class for parsing `t` tags
  class Text < OOXMLDocumentObject
    # @return [Symbol] space values
    attr_reader :space
    # @return [String] text value
    attr_reader :text

    # Parse Text object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Text] result of parsing
    def parse(node)
      @text = node.text
      node.attributes.each do |key, value|
        case key
        when 'space'
          @space = value_to_symbol(value)
        end
      end
      self
    end
  end
end
