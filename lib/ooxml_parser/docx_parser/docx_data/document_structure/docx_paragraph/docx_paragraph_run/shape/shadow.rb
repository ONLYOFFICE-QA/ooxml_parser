module OoxmlParser
  class Shadow
    attr_accessor :color, :offset

    def initialize(color = nil, offset = nil)
      @color = color
      @offset = offset
    end
  end
end
