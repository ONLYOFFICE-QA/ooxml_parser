module OoxmlParser
  # Class for parsing `w:rPrDefault` tags
  class RunPropertiesDefault < OOXMLDocumentObject
    # @return [RunProperties] properties of run
    attr_accessor :run_properties

    # Parse RunPropertiesDefault object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [RunPropertiesDefault] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'rPr'
          @run_properties = RunProperties.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
