# frozen_string_literal: true

require_relative 'table_row_properties/table_row_height'
module OoxmlParser
  # Class for describing Table Row Properties
  class TableRowProperties < OOXMLDocumentObject
    # @return [TableRowHeight] Table Row Height
    attr_accessor :height
    # @return [OoxmlSize] Table cell spacing
    attr_accessor :cells_spacing
    # @return [True, False]
    # Specifies that the current row should
    # be repeated at the top each new page on which the table is displayed.
    # >ECMA-376, 3rd Edition (June, 2011), Fundamentals and Markup Language Reference 17.4.50
    attr_accessor :table_header

    # Parse Columns data
    # @param [Nokogiri::XML:Element] node with Table Row Properties data
    # @return [TableRowProperties] value of Columns data
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'trHeight'
          @height = TableRowHeight.new(parent: self).parse(node_child)
        when 'tblCellSpacing'
          @cells_spacing = OoxmlSize.new.parse(node_child)
        when 'tblHeader'
          @table_header = option_enabled?(node_child)
        end
      end
      self
    end
  end
end
