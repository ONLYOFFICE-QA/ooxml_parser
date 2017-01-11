module OoxmlParser
  # Chart Style
  class ChartStyle < OOXMLDocumentObject
    attr_accessor :style_number

    # Parse ChartStyle
    # @param [Nokogiri::XML:Node] node with Relationships
    # @return [ChartStyle] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'style'
          @style_number = node_child.attribute('val').value.to_i
        end
      end
      self
    end
  end
end
