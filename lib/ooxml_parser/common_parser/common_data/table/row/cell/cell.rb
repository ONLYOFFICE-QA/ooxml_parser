# frozen_string_literal: true

require_relative 'properties/borders'
require_relative 'cell_properties'
module OoxmlParser
  # Class for parsing `tc` tags
  class TableCell < OOXMLDocumentObject
    attr_accessor :text_body, :properties, :grid_span, :horizontal_merge, :vertical_merge, :elements

    def initialize(parent: nil)
      @elements = []
      @parent = parent
    end

    alias cell_properties properties

    # Parse TableCell object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TableCell] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'gridSpan'
          @grid_span = value.value.to_i
        when 'hMerge'
          @horizontal_merge = value.value.to_i
        when 'vMerge'
          @vertical_merge = value.value.to_i
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'txBody'
          @text_body = TextBody.new(parent: self).parse(node_child)
        when 'tcPr'
          @properties = CellProperties.new(parent: self).parse(node_child)
        when 'p'
          @elements << DocumentStructure.default_table_paragraph_style.dup.parse(node_child,
                                                                                 0,
                                                                                 DocumentStructure.default_table_run_style,
                                                                                 parent: self)
        when 'sdt'
          @elements << StructuredDocumentTag.new(parent: self).parse(node_child)
        when 'tbl'
          @elements << Table.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
