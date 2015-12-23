# Class for parsing Run Outline properties
module OoxmlParser
  class Outline < OOXMLDocumentObject
    attr_accessor :width
    attr_accessor :color_scheme

    def initialize
      @width = 0
    end

    def self.parse(node)
      outline = Outline.new
      node.attributes.each do |key, value|
        case key
        when 'w'
          outline.width = (value.value.to_f / 12_699.3).round(2)
        end
      end
      outline.color_scheme = DocxColorScheme.parse(node)
      outline
    end
  end
end
