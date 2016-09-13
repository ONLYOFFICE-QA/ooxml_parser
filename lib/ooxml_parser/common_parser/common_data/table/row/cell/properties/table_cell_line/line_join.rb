module OoxmlParser
  class LineJoin < OOXMLDocumentObject
    attr_accessor :type, :limit

    # Parse LineJoin object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [LineJoin] result of parsing
    def parse(parent_node)
      parent_node.xpath('*').each do |line_join_node|
        case line_join_node.name
        when 'round', 'bevel'
          @type = line_join_node.name.to_sym
        when 'miter'
          @type = :miter
          @limit = line_join_node.attribute('lim').value.to_f
        end
      end
      self
    end
  end
end
