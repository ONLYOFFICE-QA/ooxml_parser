require_relative 'sheet_view/pane'
module OoxmlParser
  # Class for `sheetView` data
  class SheetView < OOXMLDocumentObject
    attr_accessor :pane

    # Parse SheetView object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [SheetView] result of parsing
    def parse(node)
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
