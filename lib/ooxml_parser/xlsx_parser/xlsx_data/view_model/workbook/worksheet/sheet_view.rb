# frozen_string_literal: true

require_relative 'sheet_view/pane'
require_relative 'sheet_view/selection'
module OoxmlParser
  # Class for `sheetView` data
  class SheetView < OOXMLDocumentObject
    attr_accessor :pane
    # @return [True, False] Flag indicating whether this sheet should display gridlines.
    attr_accessor :show_gridlines
    # @return [True, False] Flag indicating whether the sheet should display row and column headings.
    attr_accessor :show_row_column_headers
    # @return [Coordinates] Reference to the top left cell
    attr_reader :top_left_cell
    # @return [Integer] Id of workbook view
    attr_reader :workbook_view_id
    # @return [Integer] Zoom scale
    attr_reader :zoom_scale
    # @return [Selection] Properties of selection
    attr_reader :selection

    def initialize(parent: nil)
      @show_gridlines = true
      @show_row_column_headers = true
      super
    end

    # Parse SheetView object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [SheetView] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'showGridLines'
          @show_gridlines = attribute_enabled?(value)
        when 'showRowColHeaders'
          @show_row_column_headers = attribute_enabled?(value)
        when 'topLeftCell'
          @top_left_cell = Coordinates.new.parse_string(value.value)
        when 'workbookViewId'
          @workbook_view_id = value.value.to_i
        when 'zoomScale'
          @zoom_scale = value.value.to_i
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'pane'
          @pane = Pane.new(parent: self).parse(node_child)
        when 'selection'
          @selection = Selection.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
