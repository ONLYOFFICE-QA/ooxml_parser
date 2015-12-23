module OoxmlParser
  class SchemeColor
    attr_accessor :value, :properties, :converted_color

    def initialize(value = nil, converted_color = 0)
      @value = value
      @converted_color = converted_color
    end

    def to_s
      @converted_color.to_s
    end
  end
end
