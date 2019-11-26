# frozen_string_literal: true

module OoxmlParser
  # Single Chart Point
  class ChartPoint
    attr_accessor :coordinates, :value

    def initialize(coordinates = nil, value = nil)
      @coordinates = coordinates
      @value = value
    end
  end
end
