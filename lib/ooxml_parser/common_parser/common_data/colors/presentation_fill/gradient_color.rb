require_relative 'gradient_color/gradient_stop'
require_relative 'gradient_color/linear_gradient'
module OoxmlParser
  # Class for parsing `gradFill` tags
  class GradientColor < OOXMLDocumentObject
    attr_accessor :gradient_stops, :path
    # @return [LinearGradient] content of Linear Gradient
    attr_accessor :linear_gradient

    def initialize(parent: nil)
      @gradient_stops = []
      @parent = parent
    end

    # Parse GradientColor object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [GradientColor] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'gsLst'
          node_child.xpath('*').each do |gradient_stop_node|
            @gradient_stops << GradientStop.new(parent: self).parse(gradient_stop_node)
          end
        when 'path'
          @path = node_child.attribute('path').value.to_sym
        when 'lin'
          @linear_gradient = LinearGradient.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
