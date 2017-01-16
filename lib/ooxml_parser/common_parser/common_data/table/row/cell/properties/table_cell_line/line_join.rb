module OoxmlParser
  class LineJoin < OOXMLDocumentObject
    attr_accessor :type, :limit

    # Parse LineJoin object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [LineJoin] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'round', 'bevel'
          @type = node_child.name.to_sym
        when 'miter'
          @type = :miter
          @limit = node_child.attribute('lim').value.to_f
        end
      end
      self
    end
  end
end
