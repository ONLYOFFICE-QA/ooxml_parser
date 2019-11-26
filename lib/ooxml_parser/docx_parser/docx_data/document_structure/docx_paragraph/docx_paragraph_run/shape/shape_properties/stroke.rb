# frozen_string_literal: true

module OoxmlParser
  # Class for working with Stroke
  class Stroke
    attr_accessor :weight, :color

    def initialize(weight = nil, color = nil)
      @weight = weight
      @color = color
    end
  end
end
