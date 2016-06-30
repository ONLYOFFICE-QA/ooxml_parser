require_relative 'docx_shape_line_path/docx_shape_line_element'
module OoxmlParser
  # Docx Shape Line Path
  class DocxShapeLinePath < OOXMLDocumentObject
    attr_accessor :width, :height, :fill, :stroke, :elements

    def initialize(elements = [])
      @elements = elements
    end

    def self.parse(path_node)
      shape_line_path = DocxShapeLinePath.new
      path_node.attributes.each do |key, value|
        case key
        when 'w'
          shape_line_path.width = value.value.to_f
        when 'h'
          shape_line_path.height = value.value.to_f
        when 'stroke'
          shape_line_path.stroke = value.value.to_f
        end
      end
      path_node.xpath('*').each do |line_path_element_node|
        shape_line_path.elements << DocxShapeLineElement.parse(line_path_element_node)
      end
      shape_line_path
    end
  end
end
