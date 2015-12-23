module OoxmlParser
  class DocumentBackground
    attr_accessor :color1, :size, :color2, :image, :type

    def initialize(color1 = nil, type = 'simple')
      @color1 = color1
      @type = type
    end
  end
end
