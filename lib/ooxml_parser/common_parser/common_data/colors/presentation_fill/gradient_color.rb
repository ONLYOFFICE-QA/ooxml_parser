require_relative 'gradient_color/gradient_stop'
require_relative 'gradient_color/linear_gradient'
module OoxmlParser
  class GradientColor
    attr_accessor :gradient_stops, :path
    # @return [LinearGradient] content of Linear Gradient
    attr_accessor :linear_gradient

    def initialize(colors = [])
      @gradient_stops = colors
    end

    def self.parse(gradient_fill_node)
      gradient_color = GradientColor.new
      gradient_fill_node.xpath('*').each do |gradient_fill_node_child|
        case gradient_fill_node_child.name
        when 'gsLst'
          gradient_fill_node_child.xpath('*').each do |gradient_stop_node|
            gradient_color.gradient_stops << GradientStop.new(parent: gradient_color).parse(gradient_stop_node)
          end
        when 'path'
          gradient_color.path = gradient_fill_node_child.attribute('path').value.to_sym
        when 'lin'
          gradient_color.linear_gradient = LinearGradient.parse(gradient_fill_node_child)
        end
      end
      gradient_color
    end
  end
end
