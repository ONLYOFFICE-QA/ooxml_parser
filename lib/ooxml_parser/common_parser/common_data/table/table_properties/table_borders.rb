module OoxmlParser
  # Class for describing Table Border Propertie
  class TableBorders < OOXMLDocumentObject
    # @return [BordersProperties] left border propertie
    attr_accessor :left
    # @return [BordersProperties] right border propertie
    attr_accessor :right
    # @return [BordersProperties] top border propertie
    attr_accessor :top
    # @return [BordersProperties] bottom border propertie
    attr_accessor :bottom

    # Parse Table Borders data
    # @param [Nokogiri::XML:Element] node with Table Borders data
    # @return [TableBorders] value of Table Borders data
    def self.parse(node)
      borders = TableBorders.new
      node.xpath('*').each do |cell_borders_node|
        case cell_borders_node.name
        when 'left'
          borders.left = BordersProperties.parse(cell_borders_node)
        when 'top'
          borders.top = BordersProperties.parse(cell_borders_node)
        when 'right'
          borders.right = BordersProperties.parse(cell_borders_node)
        when 'bottom'
          borders.bottom = BordersProperties.parse(cell_borders_node)
        end
      end
      borders
    end
  end
end
