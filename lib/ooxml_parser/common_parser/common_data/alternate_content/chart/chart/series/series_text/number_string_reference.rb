# frozen_string_literal: true

require_relative 'string_reference/number_string_cache'
module OoxmlParser
  # Class for parsing `c:strRef` object
  class NumberStringReference < OOXMLDocumentObject
    # @return [String] formula
    attr_reader :formula
    # @return [NumberStringCache] cache of string
    attr_reader :cache

    # Parse Order
    # @param [Nokogiri::XML:Node] node with Order
    # @return [Order] result of parsing
    def parse(node)
      node.xpath('*').each do |reference_child|
        case reference_child.name
        when 'f'
          @formula = reference_child.text
        when 'strCache', 'numCache'
          @cache = NumberStringCache.new(parent: self).parse(reference_child)
        end
      end
      self
    end
  end
end
