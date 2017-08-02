module OoxmlParser
  # Class for parsing `color` tag
  class OoxmlColor < OOXMLDocumentObject
    # @return [Color] value of color
    attr_reader :value
    # @return [Symbol] theme color name
    attr_reader :theme_color
    # @return [Symbol] theme shade
    attr_reader :theme_shade

    # Parse Hyperlink object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Hyperlink] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = Color.new.parse_hex_string(value.value.to_s)
        when 'themeColor'
          @theme_color = value.value.to_sym
        when 'themeShade'
          @theme_shade = Integer("0x#{value.value}")
        end
      end
    end
  end
end
