# frozen_string_literal: true

module OoxmlParser
  # Class for working with SchemeColor
  class SchemeColor
    attr_accessor :value, :properties, :converted_color

    def initialize(value = nil, converted_color = 0, parent: nil)
      @value = value
      @converted_color = converted_color
      @parent = parent
    end

    def to_s
      @converted_color.to_s
    end
  end
end
