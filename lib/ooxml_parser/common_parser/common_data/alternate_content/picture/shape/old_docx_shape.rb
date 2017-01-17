require_relative 'old_docx_shape_properties'
# Fallback DOCX shape data
module OoxmlParser
  class OldDocxShape < OOXMLDocumentObject
    attr_accessor :properties, :text_box, :fill
    # @return [FileReference] image structure
    attr_accessor :file_reference

    def parse(node)
      @properties = OldDocxShapeProperties.parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'textbox'
          @text_box = TextBox.parse_list(node_child, parent: self)
        when 'imagedata'
          @file_reference = FileReference.new(parent: self).parse(node_child)
        when 'fill'
          @fill = OldDocxShapeFill.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
