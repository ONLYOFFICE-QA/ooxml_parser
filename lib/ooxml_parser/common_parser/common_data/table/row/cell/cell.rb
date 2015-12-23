require_relative 'properties/borders'
require_relative 'cell_properties'
module OoxmlParser
  class TableCell < OOXMLDocumentObject
    attr_accessor :text_body, :properties, :grid_span, :horizontal_merge, :vertical_merge, :elements

    def initialize(text_body = nil)
      @text_body = text_body
      @elements = []
    end

    alias_method :cell_properties, :properties

    def self.parse(cell_node)
      cell = TableCell.new
      cell_node.xpath('*').each do |cell_node_child|
        case cell_node_child.name
        when 'txBody'
          cell.text_body = TextBody.parse(cell_node_child)
        when 'tcPr'
          cell.properties = CellProperties.parse(cell_node_child)
        when 'p'
          paragraph = DocxParagraph.parse(cell_node_child, 0, DocumentStructure.default_table_paragraph_style, DocumentStructure.default_table_run_style)
          cell.elements << paragraph
        when 'tbl'
          table = Table.parse(cell_node_child)
          cell.elements << table
        end
      end
      cell_node.attributes.each do |key, value|
        case key
        when 'gridSpan'
          cell.grid_span = value.value.to_i
        when 'hMerge'
          cell.horizontal_merge = value.value.to_i
        when 'vMerge'
          cell.vertical_merge = value.value.to_i
        end
      end
      cell
    end
  end
end
