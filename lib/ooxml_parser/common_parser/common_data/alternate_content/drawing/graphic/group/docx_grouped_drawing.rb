require_relative 'docx_group_element'
# Docx Groping Drawing Data
module OoxmlParser
  class DocxGroupedDrawing < OOXMLDocumentObject
    attr_accessor :elements, :properties

    def initialize(elements = [])
      @elements = elements
    end

    def self.parse(grouping_node)
      grouping = DocxGroupedDrawing.new
      grouping_node.xpath('*').each do |grouping_node_child|
        case grouping_node_child.name
        when 'grpSpPr'
          grouping.properties = DocxShapeProperties.parse(grouping_node_child)
        when 'cNvGrpSpPr'
        when 'pic'
          element = DocxGroupElement.new(:picture)
          element.object = DocxPicture.parse(grouping_node_child)
          grouping.elements << element
        when 'wsp'
          element = DocxGroupElement.new(:shape)
          element.object = DocxShape.parse(grouping_node_child)
          grouping.elements << element
        when 'grpSp'
          element = parse(grouping_node_child)
          grouping.elements << element
        end
      end
      grouping
    end
  end
end
