require_relative 'sheet_view/pane'
module OoxmlParser
  # Class for `sheetView` data
  class SheetView < OOXMLDocumentObject
    attr_accessor :pane
    # @return [True, False] Flag indicating whether this sheet should display gridlines.
    attr_accessor :show_gridlines
    # @return [True, False] Flag indicating whether the sheet should display row and column headings.
    attr_accessor :show_row_column_headers

    def initialize(parent: nil)
      @parent = parent
      @show_gridlines = true
      @show_row_column_headers = true
    end

    # Parse SheetView object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [SheetView] result of parsing
    def parse(node)
      node.attributes.each do |key, _value|
        case key
        when 'showGridLines'
          @show_gridlines = attribute_enabled?(node, key)
        when 'showRowColHeaders'
          @show_row_column_headers = attribute_enabled?(node, key)
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'pane'
          @pane = Pane.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
