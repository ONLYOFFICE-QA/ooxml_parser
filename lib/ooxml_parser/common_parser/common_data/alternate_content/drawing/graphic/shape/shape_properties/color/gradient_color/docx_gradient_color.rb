require_relative 'docx_single_gradient_color'
# docx Gradient Color
module OoxmlParser
  class DocxGradientColor < OOXMLDocumentObject
    attr_accessor :gradient_stops, :type, :path

    def initialize(colors = [])
      @gradient_stops = colors
    end

    def self.parse(gradient_fill_node)
      gradient_color = DocxGradientColor.new
      gradient_fill_node.xpath('*').each do |grad_fill_color_child|
        case grad_fill_color_child.name
        when 'gsLst'
          grad_fill_color_child.xpath('a:gs', 'xmlns:a' => 'http://schemas.openxmlformats.org/drawingml/2006/main').each do |gradient_single_color_node|
            # gradient_single_color_node.xpath('*').each do |gradient_color_node|
            single_gradient_color = DocxSingleGradientColor.new
            single_gradient_color.color = Color.parse_color_model(gradient_single_color_node)
            single_gradient_color.position = gradient_single_color_node.attribute('pos').value.to_f / 1_000.0
            gradient_color.gradient_stops << single_gradient_color
          end
        when 'path'
          gradient_color.path = grad_fill_color_child.attribute('path').value.to_sym
        end
      end
      gradient_color
    end
  end
end
