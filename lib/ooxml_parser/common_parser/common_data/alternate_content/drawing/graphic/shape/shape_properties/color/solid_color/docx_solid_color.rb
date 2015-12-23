# Docx Solid Color
module OoxmlParser
  class DocxSolidColor
    attr_accessor :color, :luminance_modulation

    def initialize(color = nil)
      @color = color
    end
  end
end
