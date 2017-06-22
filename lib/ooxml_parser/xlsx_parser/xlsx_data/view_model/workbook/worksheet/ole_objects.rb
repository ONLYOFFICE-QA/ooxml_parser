module OoxmlParser
  # Class for parsing `oleObjects`
  class OleObjects < OOXMLDocumentObject
    # Parse OleObjects object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Array] list of OleObjects
    def parse(node)
      list = []
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'AlternateContent'
          list << AlternateContent.new(parent: self).parse(node_child)
        end
      end
      list
    end
  end
end
