require_relative 'old_docx_shape_properties'
# Fallback DOCX shape data
module OoxmlParser
  class OldDocxShape < OOXMLDocumentObject
    attr_accessor :properties, :text_box, :fill
    # @return [FileReference] image structure
    attr_accessor :file_reference

    def self.parse(shape_node)
      shape = OldDocxShape.new
      shape.properties = OldDocxShapeProperties.parse(shape_node)
      shape_node.xpath('*').each do |shape_node_child|
        case shape_node_child.name
        when 'textbox'
          shape.text_box = TextBox.parse_list(shape_node_child)
        when 'imagedata'
          shape.file_reference = FileReference.parse(shape_node_child)
        when 'fill'
          shape.fill = OldDocxShapeFill.parse(shape_node_child)
        end
      end
      shape
    end
  end
end
