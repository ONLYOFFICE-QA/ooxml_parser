require_relative 'fill/pattern_fill'
module OoxmlParser
  # Parsing `fill` tag
  class Fill < OOXMLDocumentObject
    # @return [PatternFill] pattern fill
    attr_accessor :pattern_fill

    # Parse Fill data
    # @param [Nokogiri::XML:Element] node with Fill data
    # @return [Fill] value of Fill data
    def parse(node)
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
