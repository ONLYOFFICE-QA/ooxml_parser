module OoxmlParser
  class Stroke
    attr_accessor :weight, :color

    def initialize(weight = nil, color = nil)
      @weight = weight
      @color = color
    end
  end
end
