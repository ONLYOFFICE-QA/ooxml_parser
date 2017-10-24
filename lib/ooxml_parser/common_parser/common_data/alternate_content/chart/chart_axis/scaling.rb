module OoxmlParser
  # Parsing Scaling tag 'scaling'
  class Scaling < OOXMLDocumentObject
    # @return [ValuedChild] orientation value
    attr_reader :orientation

    # Parse Scaling object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Scaling] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'orientation'
          @orientation = ValuedChild.new(:symbol, parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
