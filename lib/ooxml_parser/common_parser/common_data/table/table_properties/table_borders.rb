module OoxmlParser
  # Class for describing Table Border Propertie
  class TableBorders < OOXMLDocumentObject
    # @return [BordersProperties] left border property
    attr_accessor :left
    # @return [BordersProperties] right border property
    attr_accessor :right
    # @return [BordersProperties] top border property
    attr_accessor :top
    # @return [BordersProperties] bottom border property
    attr_accessor :bottom
    # @return [BordersProperties] inside vertical property
    attr_accessor :inside_vertical
    # @return [BordersProperties] inside horizontal property
    attr_accessor :inside_horizontal

    # Parse Table Borders data
    # @param [Nokogiri::XML:Element] node with Table Borders data
    # @return [TableBorders] value of Table Borders data
    def self.parse(node)
      borders = TableBorders.new
      node.xpath('*').each do |cell_borders_node|
        case cell_borders_node.name
        when 'left'
          borders.left = BordersProperties.new(parent: borders).parse(cell_borders_node)
        when 'top'
          borders.top = BordersProperties.new(parent: borders).parse(cell_borders_node)
        when 'right'
          borders.right = BordersProperties.new(parent: borders).parse(cell_borders_node)
        when 'bottom'
          borders.bottom = BordersProperties.new(parent: borders).parse(cell_borders_node)
        when 'insideV'
          borders.inside_vertical = BordersProperties.new(parent: borders).parse(cell_borders_node)
        when 'insideH'
          borders.inside_horizontal = BordersProperties.new(parent: borders).parse(cell_borders_node)
        end
      end
      borders
    end
  end
end
