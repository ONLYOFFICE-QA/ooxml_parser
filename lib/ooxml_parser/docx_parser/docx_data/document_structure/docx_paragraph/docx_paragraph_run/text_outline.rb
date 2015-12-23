# Class for parsing text outline and its properties
module OoxmlParser
  class TextOutline < OOXMLDocumentObject
    attr_accessor :width, :color_scheme

    def initialize
      @width = 0
      @color_scheme = :none
    end

    def self.parse(node)
      text_outline = TextOutline.new

      unless node.attribute('w').nil?
        text_outline.width = (node.attribute('w').value.to_f / 12_699).round(2)
      end

      text_outline.color_scheme = DocxColorScheme.parse(node)
      text_outline
    end
  end
end
