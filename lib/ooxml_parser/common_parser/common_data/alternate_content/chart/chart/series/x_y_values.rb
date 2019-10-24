# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `c:xVal`, `c:yVal` object
  class XYValues < OOXMLDocumentObject
    # @return [NumberReference] number reference
    attr_reader :number_reference

    # Parse XYValues
    # @param [Nokogiri::XML:Node] node with XYValues
    # @return [XYValues] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'numRef'
          @number_reference = NumberStringReference.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
