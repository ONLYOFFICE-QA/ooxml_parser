# frozen_string_literal: true

module OoxmlParser
  # Class for `CustomFilter` data
  class CustomFilter < OOXMLDocumentObject
    # @return [Symbol] operator for filter
    attr_reader :operator
    # @return [Integer] value of filter
    attr_reader :value

    # Parse CustomFilter data
    # @param [Nokogiri::XML:Element] node with CustomFilter data
    # @return [CustomFilter] value of CustomFilter data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'operator'
          @operator = value_to_symbol(value)
        when 'val'
          @value = value.value.to_i
        end
      end
      self
    end
  end
end
