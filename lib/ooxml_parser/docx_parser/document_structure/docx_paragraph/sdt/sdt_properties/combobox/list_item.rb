# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `listItem` tag
  class ListItem < OOXMLDocumentObject
    # @return [String] item text
    attr_reader :text
    # @return [String] item value
    attr_reader :value

    # Parse ListItem object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ListItem] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'displayText'
          @text = value.value
        when 'value'
          @value = value.value
        end
      end
      self
    end
  end
end
