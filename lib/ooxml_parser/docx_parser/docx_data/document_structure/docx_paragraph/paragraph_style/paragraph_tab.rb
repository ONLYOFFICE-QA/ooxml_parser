module OoxmlParser
  # Class for storing `w:tab` data
  class ParagraphTab < OOXMLDocumentObject
    # @return [Symbol] Specifies the style of the tab.
    attr_accessor :value
    # @return [OOxmlSize] Specifies the position of the tab stop.
    attr_accessor :position

    alias align value

    # Parse ParagraphTab object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ParagraphTab] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = value.value.to_sym
        when 'pos'
          @position = OoxmlSize.new(value.value.to_f)
        end
      end
      self
    end
  end
end
