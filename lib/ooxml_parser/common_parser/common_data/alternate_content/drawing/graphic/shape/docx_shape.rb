require_relative 'docx_shape/ooxml_text_box'
require_relative 'shape_properties/docx_shape_properties'
require_relative 'shape_body_properties/ooxml_shape_body_properties'
# Docx Shape Data
module OoxmlParser
  class DocxShape < OOXMLDocumentObject
    attr_accessor :non_visual_properties, :properties, :style, :body_properties, :text_body

    def self.parse(shape_node, parent: nil)
      shape = DocxShape.new
      shape.parent = parent
      shape_node.xpath('*').each do |shape_node_child|
        case shape_node_child.name
        when 'spPr'
          shape.properties = DocxShapeProperties.parse(shape_node_child)
        when 'txbx'
          shape.text_body = OOXMLTextBox.parse(shape_node_child, parent: shape)
        when 'txBody'
          shape.text_body = TextBody.parse(shape_node_child)
        when 'bodyPr'
          shape.body_properties = OOXMLShapeBodyProperties.parse(shape_node_child)
        end
      end
      shape
    end
  end
end
