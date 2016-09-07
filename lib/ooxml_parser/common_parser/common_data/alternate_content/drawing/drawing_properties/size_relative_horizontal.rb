require_relative 'size_relative_horizontal/picture_width'
module OoxmlParser
  # Class for parsing `wp14:sizeRelH` object
  class SizeRelativeHorizontal < OOXMLDocumentObject
    # @return [Symbol] type from which is relative
    attr_accessor :relative_from
    # @return [PictureWidth] width class
    attr_accessor :width

    # Parse SizeRelativeHeight
    # @param [Nokogiri::XML:Node] node with SizeRelativeHeight
    # @return [SizeRelativeHorizontal] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'relativeFrom'
          @relative_from = value_to_symbol(value)
        end
      end

      node.xpath('*').each do |preset_geometry_child|
        case preset_geometry_child.name
        when 'pctWidth'
          @width = PictureWidth.parse(preset_geometry_child)
        end
      end

      self
    end
  end
end
