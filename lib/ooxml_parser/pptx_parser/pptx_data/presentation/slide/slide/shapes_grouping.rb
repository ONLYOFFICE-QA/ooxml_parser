module OoxmlParser
  class ShapesGrouping
    attr_accessor :elements, :properties

    def initialize(elements = [])
      @elements = elements
    end

    def self.parse(grouping_node)
      grouping = ShapesGrouping.new
      grouping_node.xpath('*').each do |grouping_node_child|
        case grouping_node_child.name
        when 'grpSpPr'
          grouping.properties = DocxShapeProperties.new(parent: grouping).parse(grouping_node_child)
        when 'pic'
          grouping.elements << DocxPicture.parse(grouping_node_child)
        when 'sp'
          grouping.elements << PresentationShape.parse(grouping_node_child).dup
        when 'grpSp'
          grouping.elements << parse(grouping_node_child)
        when 'graphicFrame'
          grouping.elements << GraphicFrame.parse(grouping_node_child)
        end
      end
      grouping
    end
  end
end
