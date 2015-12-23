module OoxmlParser
  class Background < OOXMLDocumentObject
    attr_accessor :fill, :shade_to_title

    def initialize(type = nil)
      @type = type
    end

    def self.parse(background_node)
      background = Background.new
      background_properties_node = background_node.xpath('p:bgPr').first
      if background_properties_node
        background.shade_to_title = OOXMLDocumentObject.option_enabled?(background_properties_node, 'shadeToTitle')
        background.fill = PresentationFill.parse(background_properties_node)
      end
      background
    end
  end
end
