module OoxmlParser
  class LinearGradient
    attr_accessor :angle, :scaled

    def initialize(angle = nil, scaled = nil)
      @angle = angle
      @scaled = scaled
    end
  end
end
