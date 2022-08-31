# frozen_string_literal: true

module OoxmlParser
  # Class for parsing marker layout
  class MarkerLayout < OOXMLDocumentObject
    # @return [Integer] size of marker
    attr_reader :size
    # @return [Symbol] symbol of marker
    attr_reader :symbol

    # Parse Marker Layout
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [MarkerLayout] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'size'
          @size = value.value.to_i
        when 'symbol'
          @symbol = value.value.to_sym
        end
      end
      self
    end
  end
end
