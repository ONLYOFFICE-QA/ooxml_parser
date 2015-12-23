module OoxmlParser
  class GradientStop < OOXMLDocumentObject
    attr_accessor :position, :color

    def initialize(position)
      @position = position
    end

    def self.parse(gs_node)
      gradient_stop = GradientStop.new(gs_node.attribute('pos').value.to_i / 1_000)
      gs_node.xpath('*').each { |color_node| gradient_stop.color = Color.parse_color(color_node) }
      gradient_stop
    end
  end
end
