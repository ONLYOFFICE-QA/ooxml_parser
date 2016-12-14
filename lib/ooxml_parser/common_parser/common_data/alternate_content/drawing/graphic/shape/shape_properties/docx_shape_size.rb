# Docx Shape Size
module OoxmlParser
  class DocxShapeSize < OOXMLDocumentObject
    attr_accessor :rotation, :flip_horizontal, :flip_vertical, :offset, :extent
    attr_accessor :child_offset
    attr_accessor :child_extent

    alias extents extent

    # Parse DocxShapeSize object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocxShapeSize] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'rot'
          @rotation = value.value.to_f
        when 'flipH'
          @flip_horizontal = value.value.to_f
        when 'flipV'
          @flip_vertical = value.value.to_f
        end
      end
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'off'
          @offset = OOXMLCoordinates.parse(node_child, unit: :emu)
        when 'ext'
          @extent = OOXMLCoordinates.parse(node_child, x_attr: 'cx', y_attr: 'cy', unit: :emu)
        when 'chOff'
          @child_offset = OOXMLCoordinates.parse(node_child, unit: :emu)
        when 'chExt'
          @child_extent = OOXMLCoordinates.parse(node_child, x_attr: 'cx', y_attr: 'cy', unit: :emu)
        end
      end
      self
    end
  end
end
