require_relative 'shape/old_docx_shape'
require_relative 'shape/old_docx_shape_fill'
require_relative 'group/old_docx_group'
# Fallback DOCX Picture
module OoxmlParser
  class OldDocxPicture < OOXMLDocumentObject
    attr_accessor :data, :type, :style_number

    def self.parse(picture_node, parent: nil)
      picture = OldDocxPicture.new
      picture.parent = parent
      picture_node.xpath('*').each do |picture_node_child|
        case picture_node_child.name
        when 'shape'
          picture.type = :shape
          picture.data = OldDocxShape.parse(picture_node_child, parent: picture)
        when 'group'
          picture.type = :group
          picture.data = OldDocxGroup.parse(picture_node_child)
        when 'style'
          picture.style_number = picture_node_child.attribute('val').value.to_i
        end
      end
      picture
    end
  end
end
