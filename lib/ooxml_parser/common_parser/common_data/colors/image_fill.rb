require_relative 'tile'
require_relative 'image_properties'
require_relative 'image/stretching'
module OoxmlParser
  class ImageFill < OOXMLDocumentObject
    attr_accessor :path, :stretch, :tile, :properties

    alias path_to_media_file path

    def initialize(path = '')
      @path = path
    end

    def self.parse(blip_fill_node)
      image = ImageFill.new
      blip_fill_node.xpath('*').each do |blip_fill_node_child|
        case blip_fill_node_child.name
        when 'blip'
          next unless blip_fill_node_child.attribute('embed')
          path_to_original_image = OOXMLDocumentObject.get_link_from_rels(blip_fill_node_child.attribute('embed').value)
          FileUtils.copy(dir + path_to_original_image, OOXMLDocumentObject.media_folder + File.basename(path_to_original_image))
          image.path = OOXMLDocumentObject.media_folder + File.basename(path_to_original_image)
          image.properties = ImageProperties.parse(blip_fill_node_child)
        when 'stretch'
          image.stretch = Stretching.parse(blip_fill_node_child) # !blip_fill_node_child.xpath('a:fillRect').first.nil?
        when 'tile'
          image.tile = Tile.parse(blip_fill_node_child)
        end
      end
      image
    end
  end
end
