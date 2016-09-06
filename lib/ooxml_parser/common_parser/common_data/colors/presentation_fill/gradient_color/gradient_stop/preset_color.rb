module OoxmlParser
  # Class for working with `a:prstClr` tag
  class PresetColor < OOXMLDocumentObject
    # @return [String] value of preset color
    attr_accessor :value

    # Parse PresetColor object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [PresetColor] result of parsing
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
