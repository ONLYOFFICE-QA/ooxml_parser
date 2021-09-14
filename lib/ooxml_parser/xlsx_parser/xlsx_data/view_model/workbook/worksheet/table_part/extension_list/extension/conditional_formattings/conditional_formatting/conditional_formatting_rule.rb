# frozen_string_literal: true

require_relative 'conditional_formatting_rule/differential_formatting_record'
require_relative 'conditional_formatting_rule/data_bar'
require_relative 'conditional_formatting_rule/color_scale'
require_relative 'conditional_formatting_rule/icon_set'
module OoxmlParser
  # Class for `cfRule` data
  class ConditionalFormattingRule < OOXMLDocumentObject
    # @return [Symbol] Type of rule
    attr_reader :type
    # @return [Integer] Specifies position on the list of rules
    attr_reader :priority
    # @return [String] ID of rule
    attr_reader :id
    # @return [Integer] index of format
    attr_reader :format_index
    # @return [Symbol] Specifies whether rules with lower priority should be applied over this rule
    attr_reader :stop_if_true
    # @return [Symbol] Relational operator in value rule
    attr_reader :operator
    # @return [Symbol] Specifies whether top/bottom rule highlights bottom values
    attr_reader :bottom
    # @return [Symbol] Specifies whether percent is used in top/bottom rule
    attr_reader :percent
    # @return [Integer] Number of items in top/bottom rule
    attr_reader :rank
    # @return [Integer] Number of standard deviations in above/below average rule
    attr_reader :standard_deviation
    # @return [String] Text value in text rule
    attr_reader :text
    # @return [Symbol] Time period in date rule
    attr_reader :time_period
    # @return [Array<Formula>] Formulas to determine condition
    attr_reader :formulas
    # @return [DifferentialFormattingRecord] Format
    attr_reader :rule_format
    # @return [DataBar] data bar formatting
    attr_reader :data_bar
    # @return [ColorScale] color scale formatting
    attr_reader :color_scale
    # @return [IconSet] icon set formatting
    attr_reader :icon_set

    def initialize(parent: nil)
      @formulas = []
      super
    end

    # Parse ConditionalFormattingRule data
    # @param [Nokogiri::XML:Element] node with ConditionalFormattingRule data
    # @return [ConditionalFormattingRule] value of ConditionalFormattingRule data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'type'
          @type = value.value.to_sym
        when 'priority'
          @priority = value.value.to_i
        when 'id'
          @id = value.value.to_s
        when 'dxfId'
          @format_index = value.value.to_i
        when 'stopIfTrue'
          @stop_if_true = attribute_enabled?(value)
        when 'operator'
          @operator = value.value.to_sym
        when 'bottom'
          @bottom = attribute_enabled?(value)
        when 'percent'
          @percent = attribute_enabled?(value)
        when 'rank'
          @rank = value.value.to_i
        when 'stdDev'
          @standard_deviation = value.value.to_i
        when 'text'
          @text = value.text.to_s
        when 'timePeriod'
          @time_period = value.value.to_sym
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'f'
          @formulas << Formula.new(parent: self).parse(node_child)
        when 'dxf'
          @rule_format = DifferentialFormattingRecord.new(parent: self).parse(node_child)
        when 'dataBar'
          @data_bar = DataBar.new(parent: self).parse(node_child)
        when 'colorScale'
          @color_scale = ColorScale.new(parent: self).parse(node_child)
        when 'iconSet'
          @icon_set = IconSet.new(parent: self).parse(node_child)
        end
      end
      self
    end

    # @return [nil, DifferentialFormattingRecord] format of rule
    def format
      return @rule_format if @rule_format
      return nil unless @format_index

      root_object.style_sheet.differential_formatting_records[@format_index]
    end
  end
end
