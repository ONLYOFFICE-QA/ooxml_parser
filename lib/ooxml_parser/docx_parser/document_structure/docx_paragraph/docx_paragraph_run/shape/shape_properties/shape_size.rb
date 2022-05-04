# frozen_string_literal: true

module OoxmlParser
  # Class for working with Shape Size
  class ShapeSize
    attr_accessor :width, :height

    def initialize(width = nil, height = nil)
      @width = width
      @height = height
    end
  end
end
