# Operator Data
module OoxmlParser
  class Operator
    attr_accessor :operator, :bottom_value, :top_value, :value

    def initialize(operator = nil, bottom_value = nil, top_value = nil, value = nil)
      @operator = operator
      @bottom_value = bottom_value
      @top_value = top_value
      @value = value
    end
  end
end
