module OoxmlParser
  # Class for parsing `a:latin`, `a:ea`, `a:cs` tag
  class TextFont < OOXMLDocumentObject
    # @return [String] typeface of font
    attr_reader :typeface

    # Parse TextFont object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TextFont] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'typeface'
          @typeface = value.value.to_s
        end
      end
      self
    end
  end
end
