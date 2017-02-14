module OoxmlParser
  # Class for parsing `stretch` tag
  class Stretch < OOXMLDocumentObject
    # Parse Stretch object
    # @param _node [Nokogiri::XML:Element] node to parse
    # @return [Stretch] result of parsing
    def parse(_node)
      self
    end
  end
end
