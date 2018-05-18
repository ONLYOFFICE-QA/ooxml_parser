require_relative 'chart_axis/scaling'
require_relative 'chart_axis_title'
module OoxmlParser
  # Parsing Chart axis tags 'catAx', 'valAx'
  class ChartAxis < OOXMLDocumentObject
    attr_accessor :title, :display, :position, :major_grid_lines, :minor_grid_lines
    # @return [Scaling] scaling attribute
    attr_reader :scaling
    # @return [ValuedChild] the position of the tick labels
    attr_reader :tick_label_position

    def initialize(title = ChartAxisTitle.new,
                   display = true,
                   major_grid_lines = false,
                   minor_grid_lines = false,
                   parent: nil)
      @title = title
      @display = display
      @minor_grid_lines = minor_grid_lines
      @major_grid_lines = major_grid_lines
      @parent = parent
    end

    # Parse ChartAxis object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ChartAxis] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'delete'
          @display = false if node_child.attribute('val').value == '1'
        when 'title'
          @title = ChartAxisTitle.new(parent: self).parse(node_child)
        when 'majorGridlines'
          @major_grid_lines = true
        when 'minorGridlines'
          @minor_grid_lines = true
        when 'scaling'
          @scaling = Scaling.new(parent: self).parse(node_child)
        when 'tickLblPos'
          @tick_label_position = ValuedChild.new(:symbol, parent: self).parse(node_child)
        when 'axPos'
          @position = value_to_symbol(node_child.attribute('val'))
        end
      end
      @display = false unless @title.visible?
      self
    end
  end
end
