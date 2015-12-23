# Fraction Data
module OoxmlParser
  class Fraction
    attr_accessor :numerator, :denominator

    def initialize(numerator = nil, denominator = nil)
      @numerator = numerator
      @denominator = denominator
    end
  end
end
