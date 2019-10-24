# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `wp14:sizeRelV` object
  class SizeRelativeVertical < OOXMLDocumentObject
    # @return [Symbol] type from which is relative
    attr_accessor :relative_from
    # @return [PictureHeight] heigth class
    attr_accessor :height

    # Parse SizeRelativeHeight
    # @param [Nokogiri::XML:Node] node with SizeRelativeVertical
    # @return [SizeRelativeVertical] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'relativeFrom'
          @relative_from = value_to_symbol(value)
        end
      end

      node.xpath('*').each do |preset_geometry_child|
        case preset_geometry_child.name
        when 'pctHeight'
          @height = PictureDimension.new(parent: self).parse(preset_geometry_child)
        end
      end

      self
    end
  end
end
