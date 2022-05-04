# frozen_string_literal: true

module OoxmlParser
  # Class for working with Shape Margins
  class ShapeMargins
    attr_accessor :left, :right, :top

    def initialize(left = nil, right = nil, top = nil)
      @left = left
      @right = right
      @top = top
    end
  end
end
