# frozen_string_literal: true

require_relative 'behavior/behavior'

module OoxmlParser
  # Class for data for Set Time Node
  class SetTimeNode < OOXMLDocumentObject
    attr_accessor :behavior, :to

    # Parse SetTimeNode object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [SetTimeNode] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'cBhvr'
          @behavior = Behavior.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
