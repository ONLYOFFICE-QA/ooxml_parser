module OoxmlParser
  class Tile < OOXMLDocumentObject
    attr_accessor :offset, :ratio, :flip, :align

    def initialize(offset = nil, ratio = nil)
      @offset = offset
      @ratio = ratio
    end

    def self.parse(tile_node)
      tile = Tile.new(OOXMLShift.parse(tile_node, 'tx', 'ty'), OOXMLShift.parse(tile_node, 'sx', 'sy'))
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
