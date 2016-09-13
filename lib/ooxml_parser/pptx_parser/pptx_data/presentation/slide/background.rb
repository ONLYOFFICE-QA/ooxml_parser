module OoxmlParser
  # Class for parsing `bg` tag
  class Background < OOXMLDocumentObject
    attr_accessor :fill, :shade_to_title

    def initialize(type = nil)
      @type = type
    end

    # Parse Background object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Background] result of parsing
    def parse(node)
      background_properties_node = node.xpath('p:bgPr').first
      if background_properties_node
        @shade_to_title = attribute_enabled?(background_properties_node, 'shadeToTitle')
        @fill = PresentationFill.new(parent: self).parse(background_properties_node)
      end
      self
    end
  end
end
