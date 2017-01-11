module OoxmlParser
  # Class for parsing `w:lang` object
  class Language < OOXMLDocumentObject
    # @return [String] value of language
    attr_accessor :value

    # Parse Language
    # @param [Nokogiri::XML:Node] node with Language
    # @return [Language] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = value.value.to_s
        end
      end
      self
    end
  end
end
