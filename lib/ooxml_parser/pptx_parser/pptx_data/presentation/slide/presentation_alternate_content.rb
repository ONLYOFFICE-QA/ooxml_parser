require_relative 'transition/transition'
module OoxmlParser
  class PresentationAlternateContent < OOXMLDocumentObject
    attr_accessor :transition, :timing, :elements

    def initialize(elements = [])
      @elements = elements
    end

    def self.parse(alternate_content_node)
      alternate_content = PresentationAlternateContent.new
      alternate_content_node.xpath('mc:Choice/*', 'xmlns:mc' => 'http://schemas.openxmlformats.org/markup-compatibility/2006').each do |choice_node_child|
        case choice_node_child.name
        when 'timing'
          alternate_content.timing = Timing.parse(choice_node_child)
        when 'transition'
          alternate_content.transition = Transition.new(parent: alternate_content).parse(choice_node_child)
        when 'sp'
          alternate_content.elements << DocxShape.parse(choice_node_child)
        end
      end
      alternate_content
    end
  end
end
