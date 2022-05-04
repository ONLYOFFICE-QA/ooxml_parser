# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <sharedItems> tag
  class SharedItems < OOXMLDocumentObject
    # @return [True, False] is contains_semi_mixed_types
    attr_reader :contains_semi_mixed_types
    # @return [True, False] is contains_string
    attr_reader :contains_string
    # @return [True, False] is contains_number
    attr_reader :contains_number
    # @return [True, False] is contains_integer
    attr_reader :contains_integer
    # @return [Integer] min value
    attr_reader :min_value
    # @return [Integer] max value
    attr_reader :max_value

    # Parse `<sharedItems>` tag
    # # @param [Nokogiri::XML:Element] node with WorksheetSource data
    # @return [sharedItems]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'containsSemiMixedTypes'
          @contains_semi_mixed_types = attribute_enabled?(value)
        when 'containsString'
          @contains_string = attribute_enabled?(value)
        when 'containsNumber'
          @contains_number = attribute_enabled?(value)
        when 'containsInteger'
          @contains_integer = attribute_enabled?(value)
        when 'minValue'
          @min_value = value.value.to_i
        when 'maxValue'
          @max_value = value.value.to_i
        end
      end
      self
    end
  end
end
