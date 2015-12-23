module OoxmlParser
  class PresentationPattern < OOXMLDocumentObject
    attr_accessor :preset, :foreground_color, :background_color

    def initialize(preset = nil)
      @preset = preset
    end

    def self.parse(pattern_node)
      pattern = PresentationPattern.new(pattern_node.attribute('prst').value.to_sym)
      pattern_node.xpath('*').each do |color_node|
        case color_node.name
        when 'fgClr'
          pattern.foreground_color = Color.parse_color(color_node.xpath('*').first)
        when 'bgClr'
          pattern.background_color = Color.parse_color(color_node.xpath('*').first)
        end
      end
      pattern
    end
  end
end
