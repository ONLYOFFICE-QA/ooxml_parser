module OoxmlParser
  class Column
    attr_accessor :width, :space, :separator

    def initialize(width = nil, space = nil)
      @width = width
      @space = space
    end
  end
end
