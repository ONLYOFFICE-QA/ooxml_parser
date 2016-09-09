# Docx Shape Size
module OoxmlParser
  class DocxShapeSize
    attr_accessor :rotation, :flip_horizontal, :flip_vertical, :offset, :extent

    alias extents extent

    def self.parse(xfrm_node)
      shape_size = DocxShapeSize.new
      xfrm_node.attributes.each do |key, value|
        case key
        when 'rot'
          shape_size.rotation = value.value.to_f
        when 'flipH'
          shape_size.flip_horizontal = value.value.to_f
        when 'flipV'
          shape_size.flip_vertical = value.value.to_f
        end
      end
      xfrm_node.xpath('*').each do |xfrm_node_child|
        case xfrm_node_child.name
        when 'off'
          shape_size.offset = OOXMLCoordinates.parse(xfrm_node_child, unit: :emu)
        when 'ext'
          shape_size.extent = OOXMLCoordinates.parse(xfrm_node_child, x_attr: 'cx', y_attr: 'cy', unit: :emu)
        end
      end
      shape_size
    end
  end
end
