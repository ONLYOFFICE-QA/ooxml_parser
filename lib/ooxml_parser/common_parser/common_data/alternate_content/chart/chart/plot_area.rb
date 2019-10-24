# frozen_string_literal: true

require_relative 'plot_area/common_chart_data'
module OoxmlParser
  # Parsing Plot Area tag 'plotArea'
  class PlotArea < OOXMLDocumentObject
    # @return [CommonChartData] bar chart
    attr_reader :bar_chart
    # @return [CommonChartData] line chart
    attr_reader :line_chart
    # @return [CommonChartData] area chart
    attr_reader :area_chart
    # @return [CommonChartData] bubble chart
    attr_reader :bubble_chart
    # @return [CommonChartData] pie chart
    attr_reader :pie_chart
    # @return [CommonChartData] radar chart
    attr_reader :radar_chart
    # @return [CommonChartData] stock chart
    attr_reader :stock_chart
    # @return [CommonChartData] surface 3D chart
    attr_reader :surface_3d_chart
    # @return [CommonChartData] line 3D chart
    attr_reader :line_3d_chart
    # @return [CommonChartData] bar 3D chart
    attr_reader :bar_3d_chart
    # @return [CommonChartData] pie 3D chart
    attr_reader :pie_3d_chart

    # Parse View3D object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [View3D] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'barChart'
          @bar_chart = CommonChartData.new(parent: self).parse(node_child)
        when 'lineChart'
          @line_chart = CommonChartData.new(parent: self).parse(node_child)
        when 'areaChart'
          @area_chart = CommonChartData.new(parent: self).parse(node_child)
        when 'bubbleChart'
          @bubble_chart = CommonChartData.new(parent: self).parse(node_child)
        when 'doughnutChart'
          @doughnut_chart = CommonChartData.new(parent: self).parse(node_child)
        when 'pieChart'
          @pie_chart = CommonChartData.new(parent: self).parse(node_child)
        when 'scatterChart'
          @scatter_chart = CommonChartData.new(parent: self).parse(node_child)
        when 'radarChart'
          @radar_chart = CommonChartData.new(parent: self).parse(node_child)
        when 'stockChart'
          @stock_chart = CommonChartData.new(parent: self).parse(node_child)
        when 'surface3DChart'
          @surface_3d_chart = CommonChartData.new(parent: self).parse(node_child)
        when 'line3DChart'
          @line_3d_chart = CommonChartData.new(parent: self).parse(node_child)
        when 'bar3DChart'
          @bar_3d_chart = CommonChartData.new(parent: self).parse(node_child)
        when 'pie3DChart'
          @pie_3d_chart = CommonChartData.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
