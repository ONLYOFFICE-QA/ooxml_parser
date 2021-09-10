# frozen_string_literal: true

module OoxmlParser
  # Class for `cfvo` data
  class ConditionalFormatValueObject < OOXMLDocumentObject
    # @return [Symbol] Value type
    attr_reader :type
    # @return [String] Value
    attr_reader :value
    # @return [Symbol] Specifies whether value uses greater than or equal to operator
    attr_reader :greater_or_equal
    # @return [Formula] Formula
    attr_reader :formula

    def initialize(parent: nil)
      @greater_or_equal = true
      super
    end

    # Parse ConditionalFormatValueObject data
    # @param [Nokogiri::XML:Element] node with ConditionalFormatValueObject data
    # @return [ConditionalFormatValueObject] value of ConditionalFormatValueObject data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'type'
          @type = value.value.to_sym
        when 'val'
          @value = value.value.to_s
        when 'gte'
          @greater_or_equal = attribute_enabled?(value)
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'f'
          @formula = Formula.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
