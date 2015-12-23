# Class for parsing text outline and its properties
module OoxmlParser
  class TextFill < OOXMLDocumentObject
    attr_accessor :color_scheme

    def self.parse(node)
      text_fill = TextFill.new

      text_fill.color_scheme = DocxColorScheme.parse(node)
      text_fill
    end
  end
end
