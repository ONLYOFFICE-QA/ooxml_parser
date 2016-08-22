module OoxmlParser
  # Docx Shape Line Element
  class DocxShapeLineElement
    attr_accessor :type, :points

    def initialize(points = [])
      @points = points
    end

    # Parse DocxShapeLineElement
    # @param [Nokogiri::XML:Node] node with DocxShapeLineElement
    # @return [DocxShapeLineElement] result of parsing
    def self.parse(node)
      line_element = DocxShapeLineElement.new
      case node.name
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

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'pt'
          line_element.points << OOXMLCoordinates.parse(node_child)
        end
      end
      line_element
    end
  end
end
