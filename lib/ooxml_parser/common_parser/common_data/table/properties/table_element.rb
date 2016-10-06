# Describe single table element
module OoxmlParser
  class TableElement < OOXMLDocumentObject
    attr_accessor :cell_style

    alias cell_properties cell_style

    # Parse TableElement object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TableElement] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'tcStyle', 'tcPr'
          @cell_style = CellProperties.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
