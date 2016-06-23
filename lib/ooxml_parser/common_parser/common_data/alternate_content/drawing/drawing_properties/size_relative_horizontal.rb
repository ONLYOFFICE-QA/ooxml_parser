require_relative 'size_relative_horizontal/picture_width'
module OoxmlParser
  # Class for parsing `wp14:sizeRelH` object
  class SizeRelativeHorizontal
    # @return [Symbol] type from which is relative
    attr_accessor :relative_from
    # @return [PictureWidth] width class
    attr_accessor :width

    # Parse SizeRelativeHeight
    # @param [Nokogiri::XML:Node] node with SizeRelativeHeight
    # @return [SizeRelativeHorizontal] result of parsing
    def self.parse(node)
      size_width = SizeRelativeHorizontal.new

      node.attributes.each do |key, value|
        case key
        when 'relativeFrom'
          size_width.relative_from = Alignment.parse(value)
        end
      end

      node.xpath('*').each do |preset_geometry_child|
        case preset_geometry_child.name
        when 'pctWidth'
          size_width.width = PictureWidth.parse(preset_geometry_child)
        end
      end

      size_width
    end
  end
end
