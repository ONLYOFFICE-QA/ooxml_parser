require_relative 'non_visual_shape_properties'
require_relative 'shape/text_body'
module OoxmlParser
  class PresentationShape
    attr_accessor :non_visual_properties, :shape_properties, :text_body

    def initialize(non_visual_properties = nil, shape_properties = nil, text_body = nil)
      @non_visual_properties = non_visual_properties
      @shape_properties = shape_properties
      @text_body = text_body
    end

    def self.parse(shape_node)
      shape = PresentationShape.new
      shape_node.xpath('*').each do |shape_node_child|
        case shape_node_child.name
        when 'nvSpPr'
          shape.non_visual_properties = NonVisualShapeProperties.parse(shape_node_child)
        when 'spPr'
          shape.shape_properties = DocxShapeProperties.parse(shape_node_child)
        when 'txBody'
          shape.text_body = TextBody.parse(shape_node_child)
        end
      end
      shape
    end
  end
end
