# frozen_string_literal: true

module OoxmlParser
  # Parsing `dxfs` tag
  class ConditionalRuleFormats < OOXMLDocumentObject
    # @return [Array, ConditionalRuleFormat] list of conditional rule formats
    attr_reader :conditional_rule_formats
    # @return [Integer] count of formats
    attr_reader :count

    def initialize(parent: nil)
      @conditional_rule_formats = []
      super
    end

    # @return [Array, ConditionalRuleFormat] accessor
    def [](key)
      @conditional_rule_formats[key]
    end

    # Parse ConditionalRuleFormats data
    # @param [Nokogiri::XML:Element] node with ConditionalRuleFormats data
    # @return [ConditionalRuleFormats] value of ConditionalRuleFormats data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'count'
          @count = value.value.to_i
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'dxf'
          @conditional_rule_formats << ConditionalRuleFormat.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
