# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `hlinkHover` tags
  class HyperlinkForHover < OOXMLDocumentObject
    attr_accessor :id, :action, :sound

    def initialize(id = '', action = '')
      @id = id
      @action = action
    end

    # Parse HyperlinkForHover
    # @param [Nokogiri::XML:Node] node with NumberingProperties
    # @return [HyperlinkForHover] result of parsing
    def parse(node)
      @id = node.attribute('id').value
      @action = node.attribute('action').value
      node.xpath('*').each do |hyperlink_for_hover_node_child|
        case hyperlink_for_hover_node_child.name
        when 'snd'
          @sound = Sound.new(parent: self).parse(hyperlink_for_hover_node_child)
        end
      end
      self
    end
  end
end
