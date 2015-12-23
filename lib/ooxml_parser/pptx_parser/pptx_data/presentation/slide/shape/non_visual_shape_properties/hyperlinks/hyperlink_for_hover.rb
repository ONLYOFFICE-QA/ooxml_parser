module OoxmlParser
  class HyperlinkForHover
    attr_accessor :id, :action, :sound

    def initialize(id = '', action = '')
      @id = id
      @action = action
    end

    def self.parse(hyperlink_for_hover_node)
      hyperlink_for_hover = HyperlinkForHover.new
      hyperlink_for_hover.id = hyperlink_for_hover_node.attribute('id').value
      hyperlink_for_hover.action = hyperlink_for_hover_node.attribute('action').value
      hyperlink_for_hover_node.xpath('*').each do |hyperlink_for_hover_node_child|
        case hyperlink_for_hover_node_child.name
        when 'snd'
          hyperlink_for_hover.sound = Sound.parse(hyperlink_for_hover_node_child)
        end
      end
      hyperlink_for_hover
    end
  end
end
