module OoxmlParser
  # Class for parsing `w:tblStylePr`
  class TableStyleProperties < OOXMLDocumentObject
    # @return [RunProperties] properties of run
    attr_accessor :run_properties
    # @return [CellProperties] properties of table cell
    attr_accessor :table_cell_properties
    # Parse table style property
    # @param node [Nokogiri::XML::Element] node to parse
    # @param [OoxmlParser::OOXMLDocumentObject] parent parent object
    # @return [TableStyleProperties]
    def self.parse(node, parent: nil)
      table_style_pr = TableStyleProperties.new
      table_style_pr.parent = parent
      node.xpath('*').each do |properties_child|
        case properties_child.name
        when 'rPr'
          table_style_pr.run_properties = RunProperties.parse(properties_child)
        when 'tcPr'
          table_style_pr.table_cell_properties = CellProperties.parse(properties_child)
        end
      end
      table_style_pr
    end
  end
end
