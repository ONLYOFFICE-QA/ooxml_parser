# frozen_string_literal: true

require_relative 'shape/old_docx_shape'
require_relative 'shape/old_docx_shape_fill'
require_relative 'group/old_docx_group'
module OoxmlParser
  # Fallback DOCX Picture
  class OldDocxPicture < OOXMLDocumentObject
    attr_accessor :data, :type, :style_number

    # Parse OldDocxPicture object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [OldDocxPicture] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'shape'
          @type = :shape
          @data = OldDocxShape.new(parent: self).parse(node_child)
        when 'group'
          @type = :group
          @data = OldDocxGroup.new(parent: self).parse(node_child)
        when 'style'
          @style_number = node_child.attribute('val').value.to_i
        end
      end
      self
    end
  end
end
