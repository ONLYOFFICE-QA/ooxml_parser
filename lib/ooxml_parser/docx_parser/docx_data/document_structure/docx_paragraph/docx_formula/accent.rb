# Accent Data
module OoxmlParser
  class Accent
    attr_accessor :symbol, :element

    def initialize(symbol = nil, element = nil)
      @symbol = symbol
      @element = element
    end
  end
end
