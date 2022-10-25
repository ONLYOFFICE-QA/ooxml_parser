# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `checkedState` tag
  class CheckBoxState < OOXMLDocumentObject
    # @return [Integer] number of state symbol
    attr_reader :value
    # @return [String] font name
    attr_reader :font

    # Parse CheckBoxState object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [CheckBoxState] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = value.value.to_i
        when 'font'
          @font = value.value
        end
      end
      self
    end
  end
end
