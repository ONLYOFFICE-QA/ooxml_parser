require_relative 'tile'
require_relative 'image/stretching'
module OoxmlParser
  class ImageFill < OOXMLDocumentObject
    attr_accessor :stretch, :tile, :properties
    # @return [FileReference] image structure
    attr_accessor :file_reference

    def initialize(parent: nil)
      @path = ''
      @parent = parent
    end

    # Parse ImageFill object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ImageFill] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'blip'
          @file_reference = FileReference.new(parent: self).parse(node_child)
        when 'stretch'
          @stretch = Stretching.new(parent: self).parse(node_child)
        when 'tile'
          @tile = Tile.new(OOXMLCoordinates.parse(node_child, x_attr: 'tx', y_attr: 'ty'),
                           OOXMLCoordinates.parse(node_child, x_attr: 'sx', y_attr: 'sy'),
                           parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
