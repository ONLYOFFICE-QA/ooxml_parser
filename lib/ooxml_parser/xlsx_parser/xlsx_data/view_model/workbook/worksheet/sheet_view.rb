require_relative 'sheet_view/pane'
# Class for Sheet View
module OoxmlParser
  class SheetView
    attr_accessor :pane

    def self.parse(sheet_view_node)
      sheet_view = SheetView.new
      sheet_view_node.xpath('*').each do |sheet_view_node_child|
        case sheet_view_node_child.name
        when 'pane'
          sheet_view.pane = Pane.parse(sheet_view_node_child)
        end
      end
      sheet_view
    end
  end
end
