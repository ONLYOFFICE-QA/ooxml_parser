# frozen_string_literal: true

require_relative 'conditional_formattings/conditional_formatting'
module OoxmlParser
  # Class for `conditionalFormattings` data
  class ConditionalFormattings < OOXMLDocumentObject
    # @return [Array<ConditionalFormatting>] list of conditional formattings
    attr_reader :conditional_formattings_list

    def initialize(parent: nil)
      @conditional_formattings_list = []
      super
    end

    # @return [Array, ConditionalFormatting] accessor
    def [](key)
      @conditional_formattings_list[key]
    end

    # Parse ConditionalFormattings data
    # @param [Nokogiri::XML:Element] node with ConditionalFormattings data
    # @return [ConditionalFormattings] value of ConditionalFormattings data
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'conditionalFormatting'
          @conditional_formattings_list << ConditionalFormatting.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
