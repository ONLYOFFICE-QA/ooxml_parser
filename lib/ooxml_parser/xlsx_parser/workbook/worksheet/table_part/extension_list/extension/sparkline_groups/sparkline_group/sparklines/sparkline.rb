# frozen_string_literal: true

module OoxmlParser
  # Class for `sparkline` data
  class Sparkline < OOXMLDocumentObject
    # @return [String] reference to source range of sparkline
    attr_reader :source_reference
    # @return [String] reference to destination range of sparkline
    attr_reader :destination_reference

    # Parse Sparkline
    # @param [Nokogiri::XML:Node] node with Sparkline
    # @return [Sparkline] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'f'
          @source_reference = node_child.text
        when 'sqref'
          @destination_reference = node_child.text
        end
      end
      self
    end
  end
end
