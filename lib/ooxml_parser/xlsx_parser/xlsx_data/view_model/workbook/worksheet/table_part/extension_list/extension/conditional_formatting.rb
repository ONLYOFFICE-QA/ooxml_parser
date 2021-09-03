# frozen_string_literal: true

require_relative 'conditional_formatting/conditional_formatting_rule'
module OoxmlParser
  # Class for `conditionalFormatting` data
  class ConditionalFormatting < OOXMLDocumentObject
    # @return [ConditionalFormattingRule] conditional formatting rule
    attr_reader :rule
    # @return [String] Ranges to which conditional formatting is applied
    attr_reader :reference_sequence

    # Parse ConditionalFormatting data
    # @param [Nokogiri::XML:Element] node with ConditionalFormatting data
    # @return [ConditionalFormatting] value of ConditionalFormatting data
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'cfRule'
          @rule = ConditionalFormattingRule.new(parent: self).parse(node_child)
        when 'sqref'
          @reference_sequence = node_child.text
        end
      end
      self
    end
  end
end
