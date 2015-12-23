# Chart Axis Title
module OoxmlParser
  class ChartAxisTitle < OOXMLDocumentObject
    attr_accessor :layout, :overlay, :elements

    def initialize(layout = nil, overlay = nil, elements = [])
      @layout = layout
      @overlay = overlay
      @elements = elements
    end

    def nil?
      @layout.nil? && @overlay.nil? && @elements.empty?
    end

    def self.parse(axis_title_node)
      title = ChartAxisTitle.new
      axis_title_node.xpath('*').each do |title_node_child|
        case title_node_child.name
        when 'tx'
          title_node_child.xpath('c:rich/*').each do |rich_node_child|
            case rich_node_child.name
            when 'p'
              Presentation.current_font_style = FontStyle.new(true) # Default font style for chart title always bold
              title.elements << Paragraph.parse(rich_node_child)
              Presentation.current_font_style = FontStyle.new
            end
          end
        when 'layout'
          title.layout = OOXMLDocumentObject.option_enabled?(title_node_child)
        when 'overlay'
          title.overlay = OOXMLDocumentObject.option_enabled?(title_node_child)
        end
      end
      title
    end
  end
end
