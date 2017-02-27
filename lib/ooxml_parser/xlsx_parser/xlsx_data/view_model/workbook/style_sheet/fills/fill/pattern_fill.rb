module OoxmlParser
  # Parsing `patternFill` tag
  class PatternFill < OOXMLDocumentObject
    # @return [String] pattern type
    attr_accessor :pattern_type
    # @return [Color] foreground color
    attr_accessor :foreground_color
    # @return [Color] background color
    attr_accessor :background_color

    # Parse PatternFill data
    # @param [Nokogiri::XML:Element] node with PatternFill data
    # @return [PatternFill] value of PatternFill data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'patternType'
          @pattern_type = value.value.to_sym
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'fgColor'
          @foreground_color = Color.parse_color_tag(node_child)
        when 'bgColor'
          @background_color = Color.parse_color_tag(node_child)
        end
      end
      self
    end
  end
end
