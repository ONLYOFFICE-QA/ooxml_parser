module OoxmlParser
  class ShapesGrouping
    attr_accessor :elements, :properties

    def initialize(elements = [])
      @elements = elements
    end
  end
end
