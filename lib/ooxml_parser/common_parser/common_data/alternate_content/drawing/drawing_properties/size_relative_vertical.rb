require_relative 'size_relative_vertical/picture_height'
module OoxmlParser
  # Class for parsing `wp14:sizeRelV` object
  class SizeRelativeVertical
    # @return [Symbol] type from which is relative
    attr_accessor :relative_from
    # @return [PictureHeight] heigth class
    attr_accessor :height

    # Parse SizeRelativeHeight
    # @param [Nokogiri::XML:Node] node with SizeRelativeVertical
    # @return [SizeRelativeVertical] result of parsing
    def self.parse(node)
      size_height = SizeRelativeVertical.new

      node.attributes.each do |key, value|
        case key
        when 'relativeFrom'
          size_height.relative_from = Alignment.parse(value)
        end
      end

      node.xpath('*').each do |preset_geometry_child|
        case preset_geometry_child.name
        when 'pctHeight'
          size_height.height = PictureWidth.parse(preset_geometry_child)
        end
      end

      size_height
    end
  end
end
