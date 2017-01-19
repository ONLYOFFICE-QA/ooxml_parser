module OoxmlParser
  # Class for parsing `pattFill` tag
  class PresentationPattern < OOXMLDocumentObject
    attr_accessor :preset, :foreground_color, :background_color

    # Parse PresentationPattern object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [PresentationPattern] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'prst'
          @preset = value.value.to_sym
        end
      end

      node.xpath('*').each do |color_node|
        case color_node.name
        when 'fgClr'
          @foreground_color = Color.parse_color(color_node.xpath('*').first)
        when 'bgClr'
          @background_color = Color.parse_color(color_node.xpath('*').first)
        end
      end
      self
    end
  end
end
