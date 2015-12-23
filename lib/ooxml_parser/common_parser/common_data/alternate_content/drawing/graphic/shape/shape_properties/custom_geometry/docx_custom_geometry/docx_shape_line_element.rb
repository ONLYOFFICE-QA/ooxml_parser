# Docx Shape Line Element
module OoxmlParser
  class DocxShapeLineElement
    attr_accessor :type, :points

    def initialize(points = [])
      @points = points
    end
  end
end
