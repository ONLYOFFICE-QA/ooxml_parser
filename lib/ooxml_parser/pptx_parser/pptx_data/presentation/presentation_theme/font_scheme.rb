module OoxmlParser
  # Class for parsing `fontScheme` tag
  class FontScheme < OOXMLDocumentObject
    attr_reader :name

    # Parse FontScheme object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [FontScheme] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'name'
          @name = value.value.to_s
        end
      end
      self
    end
  end
end
