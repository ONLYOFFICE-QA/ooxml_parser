# Docx Shape Line Path
module OoxmlParser
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
        line_element = DocxShapeLineElement.new
        case line_path_element_node.name
        when 'moveTo'
          line_element.type = :move
        when 'lnTo'
          line_element.type = :line
        when 'arcTo'
          line_element.type = :arc
        when 'cubicBezTo'
          line_element.type = :cubic_bezier
        when 'quadBezTo'
          line_element.type = :quadratic_bezier
        when 'close'
          line_element.type = :close
        end
        line_path_element_node.xpath('a:pt', 'xmlns:a' => 'http://schemas.openxmlformats.org/drawingml/2006/main').each { |point_node| line_element.points << OOXMLCoordinates.parse(point_node, delimiter: 1) }
        shape_line_path.elements << line_element
      end
      shape_line_path
    end
  end
end
