require_relative 'old_docx_shape_properties'
# Fallback DOCX shape data
module OoxmlParser
  class OldDocxShape < OOXMLDocumentObject
    attr_accessor :properties, :text_box, :image, :fill

    def self.parse(shape_node)
      shape = OldDocxShape.new
      shape.properties = OldDocxShapeProperties.parse(shape_node)
      shape_node.xpath('*').each do |shape_node_child|
        case shape_node_child.name
        when 'textbox'
          shape.text_box = TextBox.parse_list(shape_node_child)
        when 'imagedata'
          path_to_image = OOXMLDocumentObject.copy_media_file("#{OOXMLDocumentObject.root_subfolder}/#{get_link_from_rels(shape_node_child.attribute('id').value, OOXMLDocumentObject.path_to_folder)}")
          shape.image = path_to_image
        when 'fill'
          shape.fill = OldDocxShapeFill.parse(shape_node_child)
        end
      end
      shape
    end
  end
end
