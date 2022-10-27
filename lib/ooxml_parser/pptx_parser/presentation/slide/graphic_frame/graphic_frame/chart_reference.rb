# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `c:chart` object
  class ChartReference < OOXMLDocumentObject
    # @return [String] id of the chart
    attr_reader :id

    # Parse ChartReference
    # @param [Nokogiri::XML:Node] node with ChartReference
    # @return [ChartReference] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'id'
          @id = value.value.to_s
        end
      end
      self
    end
  end
end
