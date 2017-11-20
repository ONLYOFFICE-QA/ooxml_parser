require_relative 'fill/pattern_fill'
module OoxmlParser
  # Parsing `fill` tag
  class Fill < OOXMLDocumentObject
    # @return [PatternFill] pattern fill
    attr_accessor :pattern_fill
    # @return [Color] second color
    attr_reader :color2
    # @return [String] id of file
    attr_reader :id
    # @return [FileReference] file of fill
    attr_reader :file
    # @return [Symbol] value
    attr_reader :value

    # Parse Fill data
    # @param [Nokogiri::XML:Element] node with Fill data
    # @return [Fill] value of Fill data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'color2'
          @color2 = Color.new(parent: self).parse_hex_string(value.value.split(' ').first.delete('#'))
        when 'id'
          @id = value.value.to_s
          @file = FileReference.new(parent: self).parse(node)
        when 'type'
          @type = value_to_symbol(value)
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'patternFill'
          @pattern_fill = PatternFill.new(parent: self).parse(node_child)
        end
      end
      self
    end

    def to_color
      pattern_fill.foreground_color
    end
  end
end
