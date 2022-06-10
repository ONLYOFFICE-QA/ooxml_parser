# frozen_string_literal: true

require_relative 'sparkline_group/sparklines'
module OoxmlParser
  # Class for `sparklineGroup` data
  class SparklineGroup < OOXMLDocumentObject
    # @return [Symbol] display empty cells as
    attr_reader :display_empty_cells_as
    # @return [True, False] display empty cells as
    attr_reader :display_hidden
    # @return [True, False] display x axis
    attr_reader :display_x_axis
    # @return [True, False] right to left
    attr_reader :right_to_left
    # @return [True, False] minimal axis type
    attr_reader :min_axis_type
    # @return [True, False] maximal axis type
    attr_reader :max_axis_type
    # @return [Float] manual minimum value
    attr_reader :manual_min
    # @return [[Float] manual maximum value
    attr_reader :manual_max
    # @return [Symbol] type of group
    attr_reader :type
    # @return [OoxmlSize] line weight
    attr_reader :line_weight
    # @return [True, False] show high point
    attr_reader :high_point
    # @return [True, False] show low point
    attr_reader :low_point
    # @return [True, False] show first point
    attr_reader :first_point
    # @return [True, False] show last point
    attr_reader :last_point
    # @return [True, False] show negative point
    attr_reader :negative_point
    # @return [True, False] show markers
    attr_reader :markers
    # @return [OoxmlColor] color of series
    attr_reader :color_series
    # @return [OoxmlColor] high points color
    attr_reader :color_high
    # @return [OoxmlColor] low points color
    attr_reader :color_low
    # @return [OoxmlColor] first points color
    attr_reader :color_first
    # @return [OoxmlColor] last points color
    attr_reader :color_last
    # @return [OoxmlColor] negative points color
    attr_reader :color_negative
    # @return [OoxmlColor] markers color
    attr_reader :color_markers
    # @return [Sparklines] sparklines
    attr_reader :sparklines

    # Parse SparklineGroup
    # @param [Nokogiri::XML:Node] node with SparklineGroup
    # @return [SparklineGroup] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'type'
          @type = value_to_symbol(value)
        when 'displayEmptyCellsAs'
          @display_empty_cells_as = value_to_symbol(value)
        when 'displayHidden'
          @display_hidden = attribute_enabled?(value)
        when 'displayXAxis'
          @display_x_axis = attribute_enabled?(value)
        when 'rightToLeft'
          @right_to_left = attribute_enabled?(value)
        when 'minAxisType'
          @min_axis_type = value_to_symbol(value)
        when 'maxAxisType'
          @max_axis_type = value_to_symbol(value)
        when 'manualMin'
          @manual_min = value.value.to_f
        when 'manualMax'
          @manual_max = value.value.to_f
        when 'lineWeight'
          @line_weight = OoxmlSize.new(value.value.to_f, :point)
        when 'high'
          @high_point = attribute_enabled?(value)
        when 'low'
          @low_point = attribute_enabled?(value)
        when 'first'
          @first_point = attribute_enabled?(value)
        when 'last'
          @last_point = attribute_enabled?(value)
        when 'negative'
          @negative_point = attribute_enabled?(value)
        when 'markers'
          @markers = attribute_enabled?(value)
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'colorSeries'
          @color_series = OoxmlColor.new(parent: self).parse(node_child)
        when 'colorHigh'
          @color_high = OoxmlColor.new(parent: self).parse(node_child)
        when 'colorLow'
          @color_low = OoxmlColor.new(parent: self).parse(node_child)
        when 'colorFirst'
          @color_first = OoxmlColor.new(parent: self).parse(node_child)
        when 'colorLast'
          @color_last = OoxmlColor.new(parent: self).parse(node_child)
        when 'colorNegative'
          @color_negative = OoxmlColor.new(parent: self).parse(node_child)
        when 'colorMarkers'
          @color_markers = OoxmlColor.new(parent: self).parse(node_child)
        when 'sparklines'
          @sparklines = Sparklines.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
