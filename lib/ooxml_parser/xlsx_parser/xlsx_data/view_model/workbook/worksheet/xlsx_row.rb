# frozen_string_literal: true

require_relative 'xlsx_row/xlsx_cell'
module OoxmlParser
  # Single Row of XLSX
  class XlsxRow < OOXMLDocumentObject
    attr_accessor :cells, :height, :style, :hidden
    # @return [True, False] true if the row height has been manually set.
    attr_accessor :custom_height
    # @return [Integer] Indicates to which row in the sheet this <row> definition corresponds.
    attr_accessor :index

    def initialize(parent: nil)
      @cells = []
      super
    end

    # Parse XlsxRow object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [XlsxRow] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'customHeight'
          @custom_height = option_enabled?(node, 'customHeight')
        when 'ht'
          @height = OoxmlSize.new(value.value.to_f, :point)
        when 'hidden'
          @hidden = option_enabled?(node, 'hidden')
        when 'r'
          @index = value.value.to_i
        end
      end
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'c'
          @cells[Coordinates.parse_coordinates_from_string(node_child.attribute('r').value.to_s).column_number.to_i - 1] = XlsxCell.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
