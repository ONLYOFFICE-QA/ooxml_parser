module OoxmlParser
  # Class for storing `w:tab`, `a:tab` data
  class Tab < OOXMLDocumentObject
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
        when 'algn', 'val'
          @value = value_to_symbol(value)
        when 'pos'
          @position = OoxmlSize.new(value.value.to_f, position_unit(node))
        end
      end
      self
    end

    private

    # @param node [Nokogiri::XML:Element] node to determine size
    # @return [Symbol] type of size unit
    def position_unit(node)
      return :emu if node.namespace.prefix == 'a'
      :dxa
    end
  end
end
