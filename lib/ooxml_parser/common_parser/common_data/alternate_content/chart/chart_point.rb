# Single Chart Point
module OoxmlParser
  class ChartPoint
    attr_accessor :coordinates, :value

    def initialize(coordinates = nil, value = nil)
      @coordinates = coordinates
      @value = value
    end
  end
end
