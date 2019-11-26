# frozen_string_literal: true

module OoxmlParser
  # LineJoin data
  class LineJoin < OOXMLDocumentObject
    attr_accessor :type

    # Parse LineJoin object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [LineJoin] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'round', 'bevel'
          @type = node_child.name.to_sym
        end
      end
      self
    end
  end
end
