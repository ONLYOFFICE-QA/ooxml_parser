# frozen_string_literal: true

require_relative 'series_text/number_string_reference'
module OoxmlParser
  # Class for parsing `c:tx` object
  class SeriesText < OOXMLDocumentObject
    # @return [NumberStringReference] String reference of series
    attr_accessor :reference

    # Parse Order
    # @param [Nokogiri::XML:Node] node with Order
    # @return [Order] result of parsing
    def parse(node)
      node.xpath('*').each do |series_child|
        case series_child.name
        when 'strRef', 'numRef'
          @reference = NumberStringReference.new(parent: self).parse(series_child)
        end
      end
      self
    end
  end
end
