# frozen_string_literal: true

module OoxmlParser
  # Chart Axis Title `title` node
  class ChartAxisTitle < OOXMLDocumentObject
    attr_accessor :layout, :overlay, :elements

    def initialize(parent: nil)
      @elements = []
      super
    end

    # @return [Boolean] if chart title is visible
    def visible?
      @layout || @overlay || !@elements.empty?
    end

    # Parse ChartAxisTitle object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ChartAxisTitle] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'tx'
          node_child.xpath('c:rich/*').each do |rich_node_child|
            case rich_node_child.name
            when 'p'
              root_object.default_font_style = FontStyle.new(true) # Default font style for chart title always bold
              @elements << Paragraph.new(parent: self).parse(rich_node_child)
              root_object.default_font_style = FontStyle.new
            end
          end
        when 'layout'
          @layout = option_enabled?(node_child)
        when 'overlay'
          @overlay = option_enabled?(node_child)
        end
      end
      self
    end
  end
end
