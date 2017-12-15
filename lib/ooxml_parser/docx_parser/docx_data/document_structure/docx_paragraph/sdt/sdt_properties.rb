module OoxmlParser
  # Class for parsing `w:sdtPr` tags
  class SDTProperties < OOXMLDocumentObject
    # @return [ValuedChild] tag value
    attr_reader :tag
    # @return [ValuedChild] Locking Setting
    attr_reader :lock

    # Parse SDTProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [SDTProperties] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'tag'
          @tag = ValuedChild.new(:string, parent: self).parse(node_child)
        when 'lock'
          @lock = ValuedChild.new(:symbol, parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
