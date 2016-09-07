module OoxmlParser
  # Class for parsing `tile` tag
  class Tile < OOXMLDocumentObject
    attr_accessor :offset, :ratio, :flip, :align

    def initialize(offset = nil, ratio = nil, parent: nil)
      @offset = offset
      @ratio = ratio
      @parent = parent
    end

    # Parse Tile object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Tile] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'algn'
          @align = value_to_symbol(value)
        end
      end
      self
    end
  end
end
