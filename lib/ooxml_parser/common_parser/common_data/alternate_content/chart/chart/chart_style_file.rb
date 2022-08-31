# frozen_string_literal: true

require_relative 'chart_style_file/chart_style_entry'
module OoxmlParser
  # Class for parsing Chart Style data from file
  class ChartStyleFile < OOXMLDocumentObject
    # @return [ChartStyleEntry] axis title entry
    attr_reader :axis_title
    # @return [ChartStyleEntry] axis category entry
    attr_reader :category_axis
    # @return [ChartStyleEntry] chart area entry
    attr_reader :chart_area
    # @return [ChartStyleEntry] data label entry
    attr_reader :data_label
    # @return [ChartStyleEntry] data label entry
    attr_reader :data_label_callout
    # @return [ChartStyleEntry] data point entry
    attr_reader :data_point
    # @return [ChartStyleEntry] data point 3d entry
    attr_reader :data_point_3d
    # @return [ChartStyleEntry] data point line entry
    attr_reader :data_point_line
    # @return [ChartStyleEntry] data point marker entry
    attr_reader :data_point_marker
    # @return [ChartStyleEntry] data point wireframe entry
    attr_reader :data_point_wireframe
    # @return [ChartStyleEntry] data table entry
    attr_reader :data_table
    # @return [ChartStyleEntry] down bar entry
    attr_reader :down_bar
    # @return [ChartStyleEntry] drop line entry
    attr_reader :drop_line
    # @return [ChartStyleEntry] error bar entry
    attr_reader :error_bar
    # @return [ChartStyleEntry] floor entry
    attr_reader :floor
    # @return [ChartStyleEntry] gridline major entry
    attr_reader :gridline_major
    # @return [ChartStyleEntry] gridline minor entry
    attr_reader :gridline_minor
    # @return [ChartStyleEntry] high low line entry
    attr_reader :high_low_line
    # @return [ChartStyleEntry] leader line entry
    attr_reader :leader_line
    # @return [ChartStyleEntry] legend entry
    attr_reader :legend
    # @return [ChartStyleEntry] plot area entry
    attr_reader :plot_area
    # @return [ChartStyleEntry] plot area 3d entry
    attr_reader :plot_area_3d
    # @return [ChartStyleEntry] series axis entry
    attr_reader :series_axis
    # @return [ChartStyleEntry] series line entry
    attr_reader :series_line
    # @return [ChartStyleEntry] title entry
    attr_reader :title
    # @return [ChartStyleEntry] trend line entry
    attr_reader :trend_line
    # @return [ChartStyleEntry] trend line label entry
    attr_reader :trend_line_label
    # @return [ChartStyleEntry] up bar entry
    attr_reader :up_bar
    # @return [ChartStyleEntry] value axis entry
    attr_reader :value_axis
    # @return [ChartStyleEntry] wall entry
    attr_reader :wall

    # Parse Chart style file
    # @return [ChartStyle] result of parsing
    def parse(file)
      xml = parse_xml(file)
      xml.xpath('cs:chartStyle/*').each do |chart_node|
        case chart_node.name
        when 'axisTitle'
          @axis_title = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'categoryAxis'
          @category_axis = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'chartArea'
          @chart_area = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'dataLabel'
          @data_label = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'dataLabelCallout'
          @data_label_callout = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'dataPoint'
          @data_point = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'dataPoint3D'
          @data_point_3d = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'dataPointLine'
          @data_point_line = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'dataPointMarker'
          @data_point_marker = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'dataPointMarkerLayout'
          # TODO
        when 'dataPointWireframe'
          @data_point_wireframe = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'dataTable'
          @data_table = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'downBar'
          @down_bar = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'dropLine'
          @drop_line = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'errorBar'
          @error_bar = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'extLst'
          # TODO
        when 'floor'
          @floor = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'gridlineMajor'
          @gridline_major = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'gridlineMinor'
          @gridline_minor = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'hiLoLine'
          @high_low_line = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'leaderLine'
          @leader_line = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'legend'
          @legend = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'plotArea'
          @plot_area = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'plotArea3D'
          @plot_area_3d = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'seriesAxis'
          @series_axis = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'seriesLine'
          @series_line = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'title'
          @title = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'trendline'
          @trend_line = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'trendlineLabel'
          @trend_line_label = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'upBar'
          @up_bar = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'valueAxis'
          @value_axis = ChartStyleEntry.new(parent: self).parse(chart_node)
        when 'wall'
          @wall = ChartStyleEntry.new(parent: self).parse(chart_node)
        end
      end
      self
    end
  end
end
