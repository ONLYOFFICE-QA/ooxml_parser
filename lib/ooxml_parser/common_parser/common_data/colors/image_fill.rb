require_relative 'tile'
require_relative 'image_properties'
require_relative 'image/stretching'
module OoxmlParser
  class ImageFill < OOXMLDocumentObject
    attr_accessor :stretch, :tile, :properties
    # @return [FileReference] image structure
    attr_accessor :file_reference

    def initialize(path = '')
      @path = path
    end

    def self.parse(blip_fill_node)
      image = ImageFill.new
      blip_fill_node.xpath('*').each do |blip_fill_node_child|
        case blip_fill_node_child.name
        when 'blip'
          image.file_reference = FileReference.new(parent: image).parse(blip_fill_node_child)
          image.properties = ImageProperties.parse(blip_fill_node_child)
        when 'stretch'
          image.stretch = Stretching.parse(blip_fill_node_child) # !blip_fill_node_child.xpath('a:fillRect').first.nil?
        when 'tile'
          image.tile = Tile.new(OOXMLCoordinates.parse(blip_fill_node_child, x_attr: 'tx', y_attr: 'ty'),
                                OOXMLCoordinates.parse(blip_fill_node_child, x_attr: 'sx', y_attr: 'sy'),
                                parent: image).parse(blip_fill_node_child)
        end
      end
      image
    end
  end
end
