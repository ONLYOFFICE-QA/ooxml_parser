# frozen_string_literal: true

require_relative 'old_docx_shape_properties'
module OoxmlParser
  # Fallback DOCX shape data
  class OldDocxShape < OOXMLDocumentObject
    attr_accessor :properties, :text_box, :fill
    # @return [FileReference] image structure
    attr_accessor :file_reference

    # Parse OldDocxShape object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [OldDocxShape] result of parsing
    def parse(node)
      @properties = OldDocxShapeProperties.new(parent: self).parse(node)
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
