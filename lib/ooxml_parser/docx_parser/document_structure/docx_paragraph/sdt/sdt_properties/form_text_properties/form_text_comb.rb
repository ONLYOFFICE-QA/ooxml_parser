# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `comb` tag
  class FormTextComb < OOXMLDocumentObject
    # @return [OoxmlSize] field width
    attr_reader :width
    # @return [Symbol] width rule
    attr_reader :width_rule

    # Parse FormTextComb object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [FormTextComb] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'width'
          @width = OoxmlSize.new(value.value.to_f)
        when 'wRule'
          @width_rule = value.value.to_sym
        end
      end
      self
    end
  end
end
