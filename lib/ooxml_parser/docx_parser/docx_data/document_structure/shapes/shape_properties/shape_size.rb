module OoxmlParser
  class ShapeSize
    attr_accessor :width, :height

    def initialize(width = nil, height = nil)
      @width = width
      @height = height
    end
  end
end
