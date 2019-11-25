# frozen_string_literal: true

module OoxmlParser
  # Class for Wheel Slide Transition `p:wheel` tag
  class Wheel < OOXMLDocumentObject
    # @return [Integer] count of spokes
    attr_reader :spokes

    # Parse Wheel object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Wheel] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'spokes'
          @spokes = value.value.to_i
        end
      end
      self
    end
  end
end
