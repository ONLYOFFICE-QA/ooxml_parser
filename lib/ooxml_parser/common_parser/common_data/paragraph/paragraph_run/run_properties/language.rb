module OoxmlParser
  # Class for parsing `w:lang` object
  class Language < OOXMLDocumentObject
    # @return [String] value of language
    attr_accessor :value

    # Parse Language
    # @param [Nokogiri::XML:Node] node with Language
    # @return [Language] result of parsing
    def self.parse(node)
      language = Language.new
      node.attributes.each do |key, value|
        case key
        when 'val'
          language.value = value.value.to_s
        end
      end
      language
    end
  end
end
