# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `a:alphaModFix` tag
  class AlphaModFix < OOXMLDocumentObject
    attr_accessor :amount

    # Parse AlphaModFix object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [AlphaModFix] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'amt'
          @amount = OoxmlSize.new(value.value.to_f, :one_1000th_percent)
        end
      end
      self
    end
  end
end
