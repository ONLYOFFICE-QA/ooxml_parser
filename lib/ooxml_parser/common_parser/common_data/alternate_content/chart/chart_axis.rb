require_relative 'chart_axis_title'
# Chart Axis
module OoxmlParser
  class ChartAxis < OOXMLDocumentObject
    attr_accessor :title, :display, :position, :major_grid_lines, :minor_grid_lines

    def initialize(title = ChartAxisTitle.new, display = true, major_grid_lines = false, minor_grid_lines = false)
      @title = title
      @display = display
      @minor_grid_lines = minor_grid_lines
      @major_grid_lines = major_grid_lines
    end

    def self.parse(axis_node)
      axis = ChartAxis.new
      axis_node.xpath('*').each do |axis_node_child|
        case axis_node_child.name
        when 'delete'
          axis.display = false if axis_node_child.attribute('val').value == '1'
        when 'title'
          axis.title = ChartAxisTitle.parse(axis_node_child)
        when 'majorGridlines'
          axis.major_grid_lines = true
        when 'minorGridlines'
          axis.minor_grid_lines = true
        when 'axPos'
          axis.position = Alignment.parse(axis_node_child.attribute('val'))
        end
      end
      axis.display = false if axis.title.nil?
      axis
    end
  end
end
