# frozen_string_literal: true

require_relative 'data_bar/conditional_format_value_object'
module OoxmlParser
  # Class for `dataBar` data
  class DataBar < OOXMLDocumentObject
    # @return [Integer] Minimal length of the data bar as a percentage of the cell width
    attr_reader :min_length
    # @return [Integer] Maximal length of the data bar as a percentage of the cell width
    attr_reader :max_length
    # @return [Symbol] Specifies whether value is shown in a cell
    attr_reader :show_value
    # @return [Symbol] Position of axis in a cell
    attr_reader :axis_position
    # @return [Symbol] Data bar direction
    attr_reader :direction
    # @return [Symbol] Specifies whether data bar fill is gradient
    attr_reader :gradient
    # @return [Symbol] Specifies whether data bar has border
    attr_reader :border
    # @return [Symbol] Specifies whether fill color for negative values is same as for positive
    attr_reader :negative_bar_same_as_positive
    # @return [Symbol] Specifies whether border color for negative values is same as for positive
    attr_reader :negative_border_same_as_positive
    # @return [Array, ConditionalFormatValueObject] minimal and maximal value
    attr_reader :values
    # @return [Color] Fill color for positive values
    attr_reader :fill_color
    # @return [Color] Fill color for negative values
    attr_reader :negative_fill_color
    # @return [Color] Border color for positive values
    attr_reader :border_color
    # @return [Color] Border color for negative values
    attr_reader :negative_border_color
    # @return [Color] Axis color
    attr_reader :axis_color

    def initialize(parent: nil)
      @values = []
      @show_value = true
      @gradient = true
      @negative_border_same_as_positive = true
      super
    end

    # Parse DataBar data
    # @param [Nokogiri::XML:Element] node with DataBar data
    # @return [DataBar] value of DataBar data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'minLength'
          @min_length = value.value.to_i
        when 'maxLength'
          @max_length = value.value.to_i
        when 'showValue'
          @show_value = attribute_enabled?(value)
        when 'axisPosition'
          @axis_position = value.value.to_sym
        when 'direction'
          @direction = value.value.to_sym
        when 'gradient'
          @gradient = attribute_enabled?(value)
        when 'border'
          @border = attribute_enabled?(value)
        when 'negativeBarColorSameAsPositive'
          @negative_bar_same_as_positive = attribute_enabled?(value)
        when 'negativeBarBorderColorSameAsPositive'
          @negative_border_same_as_positive = attribute_enabled?(value)
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'cfvo'
          @values << ConditionalFormatValueObject.new(parent: self).parse(node_child)
        when 'fillColor'
          @fill_color = OoxmlColor.new(parent: self).parse(node_child)
        when 'negativeFillColor'
          @negative_fill_color = OoxmlColor.new(parent: self).parse(node_child)
        when 'borderColor'
          @border_color = OoxmlColor.new(parent: self).parse(node_child)
        when 'negativeBorderColor'
          @negative_border_color = OoxmlColor.new(parent: self).parse(node_child)
        when 'axisColor'
          @axis_color = OoxmlColor.new(parent: self).parse(node_child)
        end
      end
      self
    end

    # @return [ConditionalFormatValueObject] minimal value
    def minimum
      values[0]
    end

    # @return [ConditionalFormatValueObject] maximal value
    def maximum
      values[1]
    end
  end
end
