# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `format` tag
  class FormTextFormat < OOXMLDocumentObject
    # @return [Symbol] format type
    attr_reader :type
    # @return [String] value for custom formats
    attr_reader :value
    # @return [String] allowed symbols
    attr_reader :symbols

    # Parse FormTextFormat object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [FormTextFormat] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'type'
          @type = value.value.to_sym
        when 'val'
          @value = value.value
        when 'symbols'
          @symbols = value.value
        end
      end
      self
    end
  end
end
