module OoxmlParser
  class Offset
    attr_accessor :x, :y

    def initialize(x = '', y = '')
      @x = x
      @y = y
    end
  end
end
