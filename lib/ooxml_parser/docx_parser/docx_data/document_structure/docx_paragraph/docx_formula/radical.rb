# Radical Data
module OoxmlParser
  class Radical
    attr_accessor :degree, :value

    def initialize(degree = 2, value = nil)
      @degree = degree
      @value = value
    end
  end
end
