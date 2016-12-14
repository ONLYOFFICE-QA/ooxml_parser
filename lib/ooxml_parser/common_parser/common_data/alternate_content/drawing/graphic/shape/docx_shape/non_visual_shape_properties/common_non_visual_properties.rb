module OoxmlParser
  # Class for `cNvPr` data
  class CommonNonVisualProperties < OOXMLDocumentObject
    attr_accessor :name, :id, :description, :on_click_hyperlink, :hyperlink_for_hover

    # Parse CommonNonVisualProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [CommonNonVisualProperties] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'name'
          @name = value.value
        when 'id'
          @id = value.value
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'hlinkClick'
          @on_click_hyperlink = Hyperlink.parse(node_child)
        when 'hlinkHover'
          @hyperlink_for_hover = HyperlinkForHover.parse(node_child)
        end
      end
      self
    end
  end
end
