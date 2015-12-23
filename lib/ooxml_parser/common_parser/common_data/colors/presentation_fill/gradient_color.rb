require_relative 'gradient_color/gradient_stop'
require_relative 'gradient_color/linear_gradient'
module OoxmlParser
  class GradientColor
    attr_accessor :gradient_stops, :path

    def initialize(colors = [])
      @gradient_stops = colors
    end

    def self.parse(gradient_fill_node)
      gradient_color = GradientColor.new
      gradient_fill_node.xpath('*').each do |gradient_fill_node_child|
        case gradient_fill_node_child.name
        when 'gsLst'
          begin
            gradient_fill_node_child.xpath('//a:gs')
            namespace = 'a'
          rescue Nokogiri::XML::XPath::SyntaxError
            namespace = 'w14'
          end
          gradient_fill_node_child.xpath("#{namespace}:gs").each do |gradient_stop_node|
            gradient_color.gradient_stops << GradientStop.parse(gradient_stop_node)
          end
        when 'path'
          gradient_color.path = gradient_fill_node_child.attribute('path').value.to_sym
        when 'lin'
          gradient_color.path = LinearGradient.new(gradient_fill_node_child.attribute('ang').value.to_f / 100_000.0,
                                                   gradient_fill_node_child.attribute('scaled').value.to_i)
        end
      end
      gradient_color
    end
  end
end
