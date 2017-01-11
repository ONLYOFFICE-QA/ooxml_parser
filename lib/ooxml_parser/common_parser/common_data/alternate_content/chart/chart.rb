require_relative 'chart_axis'
require_relative 'chart_cells_range'
require_relative 'chart_legend'
require_relative 'chart_point'
require_relative 'display_labels_properties'
require_relative 'chart/series'
require_relative 'chart/view_3d'

module OoxmlParser
  class Chart < OOXMLDocumentObject
    attr_accessor :type, :data, :grouping, :title, :legend, :display_labels, :axises, :alternate_content, :shape_properties
    # @return [Array, Series] series of chart
    attr_accessor :series
    # @return [View3D] properties of 3D view
    attr_accessor :view_3d

    def initialize
      @type = ''
      @data = []
      @grouping = nil
      @title = nil
      @legend = nil
      @axises = []
      @series = []
    end

    def parse_properties(chart_prop_node)
      chart_prop_node.xpath('*').each do |chart_props_node_child|
        case chart_props_node_child.name
        when 'grouping'
          @grouping = chart_props_node_child.attribute('val').value.to_sym
        when 'ser'
          @series << Series.new(parent: self).parse(chart_props_node_child)
          case @type
          when :point, :bubble
            val = chart_props_node_child.xpath('c:yVal')[0]
          else
            val = chart_props_node_child.xpath('c:val')[0]
          end
          next if val.xpath('c:numRef').empty?
          @data << ChartCellsRange.parse(val.xpath('c:numRef').first).dup
        when 'dLbls'
          @display_labels = DisplayLabelsProperties.new(parent: self).parse(chart_props_node_child)
        end
      end
    end

    def self.parse
      chart = Chart.new
      chart_xml = Nokogiri::XML(File.open(OOXMLDocumentObject.current_xml))
      chart_xml.xpath('*').each do |chart_node|
        case chart_node.name
        when 'chartSpace'
          chart_node.xpath('*').each do |chart_space_node_child|
            case chart_space_node_child.name
            when 'AlternateContent'
              chart.alternate_content = AlternateContent.new(parent: chart).parse(chart_space_node_child)
            when 'spPr'
              chart.shape_properties = DocxShapeProperties.new(parent: chart).parse(chart_space_node_child)
            when 'chart'
              chart_space_node_child.xpath('*').each do |chart_node_child|
                case chart_node_child.name
                when 'plotArea'
                  chart_node_child.xpath('*').each do |plot_area_node_child|
                    next unless chart.type.empty?
                    case plot_area_node_child.name
                    when 'barChart'
                      chart.type = :bar
                      bar_dir_node = plot_area_node_child.xpath('c:barDir')
                      chart.type = :column if bar_dir_node.first.attribute('val').value == 'col'
                      chart.parse_properties(plot_area_node_child)
                    when 'lineChart'
                      chart.type = :line
                      chart.parse_properties(plot_area_node_child)
                    when 'areaChart'
                      chart.type = :area
                      chart.parse_properties(plot_area_node_child)
                    when 'bubbleChart'
                      chart.type = :bubble
                      chart.parse_properties(plot_area_node_child)
                    when 'doughnutChart'
                      chart.type = :doughnut
                      chart.parse_properties(plot_area_node_child)
                    when 'pieChart'
                      chart.type = :pie
                      chart.parse_properties(plot_area_node_child)
                    when 'scatterChart'
                      chart.type = :point
                      chart.parse_properties(plot_area_node_child)
                    when 'radarChart'
                      chart.type = :radar
                      chart.parse_properties(plot_area_node_child)
                    when 'stockChart'
                      chart.type = :stock
                      chart.parse_properties(plot_area_node_child)
                    when 'surface3DChart'
                      chart.type = :surface_3d
                      chart.parse_properties(plot_area_node_child)
                    when 'line3DChart'
                      chart.type = :line_3d
                      chart.parse_properties(plot_area_node_child)
                    when 'bar3DChart'
                      chart.type = :bar_3d
                      chart.parse_properties(plot_area_node_child)
                    when 'pie3DChart'
                      chart.type = :pie_3d
                      chart.parse_properties(plot_area_node_child)
                    end
                  end
                  chart_node_child.xpath('*').each do |plot_area_node_child|
                    case plot_area_node_child.name
                    when 'catAx', 'valAx'
                      chart.axises << ChartAxis.new(parent: chart).parse(plot_area_node_child)
                    end
                  end
                when 'title'
                  chart.title = ChartAxisTitle.new(parent: chart).parse(chart_node_child)
                when 'legend'
                  chart.legend = ChartLegend.new(parent: chart).parse(chart_node_child)
                when 'view3D'
                  chart.view_3d = View3D.new(parent: chart).parse(chart_node_child)
                end
              end
            end
          end
        end
      end
      chart
    end
  end
end
