# frozen_string_literal: true

module OoxmlParser
  class LineEnd < OOXMLDocumentObject
    attr_accessor :type, :length, :width

    # Parse LineEnd object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [LineEnd] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'type'
          @type = value.value.to_sym
        when 'w'
          @width = value_to_symbol(value)
        when 'len'
          @length = value_to_symbol(value)
        end
      end
      self
    end
  end
end
