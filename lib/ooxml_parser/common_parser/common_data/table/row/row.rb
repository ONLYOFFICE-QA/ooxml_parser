# frozen_string_literal: true

require_relative 'cell/cell'
require_relative 'row/table_row_properties'
module OoxmlParser
  # Class for data of TableRow
  class TableRow < OOXMLDocumentObject
    attr_accessor :height, :cells, :table_row_properties

    def initialize(cells = [], parent: nil)
      @cells = cells
      @parent = parent
    end

    alias table_row_height height

    # Parse TableRow object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TableRow] result of parsing
    def parse(node)
      root_object.default_font_style = FontStyle.new(true) # TODO: Add correct parsing of TableStyle.xml file and use it
      node.attributes.each do |key, value|
        case key
        when 'h'
          @height = OoxmlSize.new(value.value.to_f, :emu)
        end
      end
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'trPr'
          @table_row_properties = TableRowProperties.new(parent: self).parse(node_child)
        when 'tc'
          @cells << TableCell.new(parent: self).parse(node_child)
        end
      end
      root_object.default_font_style = FontStyle.new
      self
    end
  end
end
