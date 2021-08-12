# frozen_string_literal: true

require_relative 'chart_axis'
require_relative 'chart_cells_range'
require_relative 'chart_legend'
require_relative 'chart_point'
require_relative 'display_labels_properties'
require_relative 'chart/pivot_formats'
require_relative 'chart/plot_area'
require_relative 'chart/series'
require_relative 'chart/view_3d'

module OoxmlParser
  # Class for working with Chart data
  class Chart < OOXMLDocumentObject
    attr_accessor :type, :data, :grouping, :title, :legend, :display_labels, :axises, :alternate_content, :shape_properties
    # @return [Array, ValuedChild] array axis id
    attr_reader :axis_ids
    # @return [Array, Series] series of chart
    attr_accessor :series
    # @return [PivotFormats] list of pivot formats
    attr_accessor :pivot_formats
    # @return [PlotArea] plot area data
    attr_accessor :plot_area
    # @return [True, False] This element specifies that
    # each data marker in the series has a different color.
    # ECMA-376 (5th Edition). 21.2.2.227
    attr_reader :vary_colors
    # @return [View3D] properties of 3D view
    attr_accessor :view_3d

    def initialize(parent: nil)
      @axis_ids = []
      @type = ''
      @data = []
      @grouping = nil
      @title = nil
      @legend = nil
      @axises = []
      @series = []
      super
    end

    # Parse properties of Chart
    # @param chart_prop_node [Nokogiri::XML:Element] node to parse
    # @return [void]
    def parse_properties(chart_prop_node)
      chart_prop_node.xpath('*').each do |chart_props_node_child|
        case chart_props_node_child.name
        when 'axId'
          @axis_ids << ValuedChild.new(:integer, parent: self).parse(chart_props_node_child)
        when 'grouping'
          @grouping = chart_props_node_child.attribute('val').value.to_sym
        when 'ser'
          @series << Series.new(parent: self).parse(chart_props_node_child)
          val = case @type
                when :point, :bubble
                  chart_props_node_child.xpath('c:yVal')[0]
                else
                  chart_props_node_child.xpath('c:val')[0]
                end
          next unless val
          next if val.xpath('c:numRef').empty?

          @data << ChartCellsRange.new(parent: self).parse(val.xpath('c:numRef').first).dup
        when 'dLbls'
          @display_labels = DisplayLabelsProperties.new(parent: self).parse(chart_props_node_child)
        when 'varyColors'
          @vary_colors = option_enabled?(chart_props_node_child)
        end
      end
    end

    extend Gem::Deprecate
    deprecate :data, 'series points interface', 2020, 1

    # Parse Chart data
    # @return [Chart] result of parsing
    def parse
      chart_xml = parse_xml(OOXMLDocumentObject.current_xml)
      chart_xml.xpath('*').each do |chart_node|
        case chart_node.name
        when 'chartSpace'
          chart_node.xpath('*').each do |chart_space_node_child|
            case chart_space_node_child.name
            when 'AlternateContent'
              @alternate_content = AlternateContent.new(parent: self).parse(chart_space_node_child)
            when 'spPr'
              @shape_properties = DocxShapeProperties.new(parent: self).parse(chart_space_node_child)
            when 'chart'
              chart_space_node_child.xpath('*').each do |chart_node_child|
                case chart_node_child.name
                when 'plotArea'
                  chart_node_child.xpath('*').each do |plot_area_node_child|
                    next unless type.empty?

                    case plot_area_node_child.name
                    when 'barChart'
                      @type = :bar
                      bar_dir_node = plot_area_node_child.xpath('c:barDir')
                      @type = :column if bar_dir_node.first.attribute('val').value == 'col'
                      parse_properties(plot_area_node_child)
                    when 'lineChart'
                      @type = :line
                      parse_properties(plot_area_node_child)
                    when 'areaChart'
                      @type = :area
                      parse_properties(plot_area_node_child)
                    when 'bubbleChart'
                      @type = :bubble
                      parse_properties(plot_area_node_child)
                    when 'doughnutChart'
                      @type = :doughnut
                      parse_properties(plot_area_node_child)
                    when 'pieChart'
                      @type = :pie
                      parse_properties(plot_area_node_child)
                    when 'scatterChart'
                      @type = :point
                      parse_properties(plot_area_node_child)
                    when 'radarChart'
                      @type = :radar
                      parse_properties(plot_area_node_child)
                    when 'stockChart'
                      @type = :stock
                      parse_properties(plot_area_node_child)
                    when 'surface3DChart'
                      @type = :surface_3d
                      parse_properties(plot_area_node_child)
                    when 'line3DChart'
                      @type = :line_3d
                      parse_properties(plot_area_node_child)
                    when 'bar3DChart'
                      @type = :bar_3d
                      parse_properties(plot_area_node_child)
                    when 'pie3DChart'
                      @type = :pie_3d
                      parse_properties(plot_area_node_child)
                    end
                  end
                  parse_axis(chart_node_child)
                  @plot_area = PlotArea.new(parent: self).parse(chart_node_child)
                when 'title'
                  @title = ChartAxisTitle.new(parent: self).parse(chart_node_child)
                when 'legend'
                  @legend = ChartLegend.new(parent: self).parse(chart_node_child)
                when 'view3D'
                  @view_3d = View3D.new(parent: self).parse(chart_node_child)
                when 'pivotFmts'
                  @pivot_formats = PivotFormats.new(parent: self).parse(chart_node_child)
                end
              end
            end
          end
        end
      end
      self
    end

    private

    # Perform parsing of axis info
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [void]
    def parse_axis(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'catAx', 'valAx'
          @axises << ChartAxis.new(parent: self).parse(node_child)
        end
      end
    end
  end
end
