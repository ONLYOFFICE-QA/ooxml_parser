# frozen_string_literal: true

require_relative 'conditional_formatting_rule/conditional_rule_format'
module OoxmlParser
  # Class for `cfRule` data
  class ConditionalFormattingRule < OOXMLDocumentObject
    # @return [Symbol] Type of rule
    attr_reader :type
    # @return [Integer] Specifies position on the list of rules
    attr_reader :priority
    # @return [String] ID of rule
    attr_reader :id
    # @return [Symbol] Specifies whether rules with lower priority should be applied over this rule
    attr_reader :stop_if_true
    # @return [Symbol] Relational operator in value rule
    attr_reader :operator
    # @return [Symbol] Specifies whether percent is used in top/bottom rule
    attr_reader :percent
    # @return [Integer] Number of items in top/bottom rule
    attr_reader :rank
    # @return [Integer] Number of standard deviations in above/below average rule
    attr_reader :standard_deviation
    # @return [String] text value in text rule
    attr_reader :text
    # @return [Array, Formula] Formulas to determine condition
    attr_reader :formulas
    # @return [ConditionalRuleFormat] Format
    attr_reader :format

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
        when 'stopIfTrue'
          @stop_if_true = attribute_enabled?(value)
        when 'operator'
          @operator = value.value.to_sym
        when 'percent'
          @percent = attribute_enabled?(value)
        when 'rank'
          @rank = value.value.to_i
        when 'stdDev'
          @standard_deviation = value.value.to_i
        when 'text'
          @text = value.text.to_s
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'f'
          @formulas << Formula.new(parent: self).parse(node_child)
        when 'dxf'
          @format = ConditionalRuleFormat.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
