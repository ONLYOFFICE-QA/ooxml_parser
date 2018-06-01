module OoxmlParser
  # Single author `author`
  class Author < OOXMLDocumentObject
    attr_reader :name

    # Parse Author object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Author] result of parsing
    def parse(node)
      @name = node.text
      self
    end
  end
end
