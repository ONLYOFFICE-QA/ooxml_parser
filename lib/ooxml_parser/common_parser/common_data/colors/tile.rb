module OoxmlParser
  class Tile < OOXMLDocumentObject
    attr_accessor :offset, :ratio, :flip, :align

    def initialize(offset = nil, ratio = nil)
      @offset = offset
      @ratio = ratio
    end

    def self.parse(tile_node)
      tile = Tile.new(OOXMLCoordinates.parse(tile_node, x_attr: 'tx', y_attr: 'ty'), OOXMLCoordinates.parse(tile_node, x_attr: 'sx', y_attr: 'sy'))
      tile_node.attributes.each do |key, value|
        case key
        when 'algn'
          tile.align = Alignment.parse(value)
        end
      end
      tile
    end
  end
end
