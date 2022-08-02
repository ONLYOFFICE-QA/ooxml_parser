# frozen_string_literal: true

module OoxmlParser
  # Single Chart Point
  class ChartPoint
    attr_accessor :coordinates, :value

    # @return [String] data format for the given point
    attr_accessor :format

    def initialize(coordinates = nil, value = nil, format = nil)
      @coordinates = coordinates
      @value = value
      @format = format
    end
  end
end
