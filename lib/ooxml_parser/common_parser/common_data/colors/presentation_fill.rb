require_relative 'image_fill'
require_relative 'presentation_fill/gradient_color'
require_relative 'presentation_fill/presentation_pattern'
module OoxmlParser
  class PresentationFill < OOXMLDocumentObject
    attr_accessor :type, :image, :color, :pattern

    def self.parse(parent_fill_node)
      fill = PresentationFill.new
      return nil if parent_fill_node.xpath('*').empty?
      parent_fill_node.xpath('*').each do |fill_node|
        case fill_node.name
        when 'gradFill'
          fill.type = :gradient
          fill.color = GradientColor.parse(fill_node)
        when 'solidFill'
          fill.type = :solid
          fill.color = Color.parse_color(fill_node.xpath('*').first)
        when 'blipFill'
          fill.type = :image
          fill.image = ImageFill.parse(fill_node)
        when 'pattFill'
          fill.type = :pattern
          fill.pattern = PresentationPattern.parse(fill_node)
        when 'noFill'
          fill.type = :noneColor
          fill.color = :none
        end
      end
      fill
    end
  end
end
