# frozen_string_literal: true

require_relative 'xlsx_row/xlsx_cell'
module OoxmlParser
  # Single Row of XLSX
  class XlsxRow < OOXMLDocumentObject
    attr_accessor :height, :style, :hidden
    # @return [True, False] true if the row height has been manually set.
    attr_accessor :custom_height
    # @return [Integer] Indicates to which row in the sheet
    #   this <row> definition corresponds.
    attr_accessor :index
    # @return [Array<Cells>] cells of row, as in xml structure
    attr_reader :cells_raw

    def initialize(parent: nil)
      @cells_raw = []
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
          @cells_raw << XlsxCell.new(parent: self).parse(node_child)
        end
      end
      self
    end

    # @return [Array<XlsxCell, nil>] list of cell in row, with nil,
    #   if cell data is not stored in xml
    def cells
      return @cells if @cells.any?

      cells_raw.each do |cell|
        @cells[cell.coordinates.column_number.to_i - 1] = cell
      end

      @cells
    end
  end
end
