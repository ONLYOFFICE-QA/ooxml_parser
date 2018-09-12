require_relative 'image_fill'
require_relative 'presentation_fill/gradient_color'
require_relative 'presentation_fill/presentation_pattern'
module OoxmlParser
  class PresentationFill < OOXMLDocumentObject
    attr_accessor :type, :image, :color, :pattern

    # Parse PresentationFill object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [PresentationFill] result of parsing
    def parse(node)
      return nil if node.xpath('*').empty?

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'gradFill'
          @type = :gradient
          @color = GradientColor.new(parent: self).parse(node_child)
        when 'solidFill'
          @type = :solid
          @color = Color.new(parent: self).parse_color(node_child.xpath('*').first)
        when 'blipFill'
          @type = :image
          @image = ImageFill.new(parent: self).parse(node_child)
        when 'pattFill'
          @type = :pattern
          @pattern = PresentationPattern.new(parent: self).parse(node_child)
        when 'noFill'
          @type = :noneColor
          @color = :none
        end
      end
      return nil if @type.nil?

      self
    end
  end
end
