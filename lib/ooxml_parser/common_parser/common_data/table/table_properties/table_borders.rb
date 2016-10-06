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
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'left'
          @left = BordersProperties.new(parent: self).parse(node_child)
        when 'top'
          @top = BordersProperties.new(parent: self).parse(node_child)
        when 'right'
          @right = BordersProperties.new(parent: self).parse(node_child)
        when 'bottom'
          @bottom = BordersProperties.new(parent: self).parse(node_child)
        when 'insideV'
          @inside_vertical = BordersProperties.new(parent: self).parse(node_child)
        when 'insideH'
          @inside_horizontal = BordersProperties.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
