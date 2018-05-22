module OoxmlParser
  # Class for parsing string index `si` tag
  class StringIndex < OOXMLDocumentObject
    # @return [String] text
    attr_reader :text

    # Parse StringIndex data
    # @param [Nokogiri::XML:Element] node with StringIndex data
    # @return [StringIndex] value of StringIndex data
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 't'
          @text = node_child.text
        end
      end
      self
    end
  end
end
