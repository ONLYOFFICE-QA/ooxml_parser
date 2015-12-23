module OoxmlParser
  class ParagraphTab < OOXMLDocumentObject
    attr_accessor :align, :position

    def initialize(align, position)
      @align = align
      @position = position
    end

    def self.parse(tabs_node)
      tabs = []
      tabs_node.xpath('a:tab').each do |tab_node|
        tabs << ParagraphTab.new(Alignment.parse(tab_node.attribute('algn')),
                                 (tab_node.attribute('pos').value.to_f / 360_000.0).round(2))
      end
      tabs
    end
  end
end
