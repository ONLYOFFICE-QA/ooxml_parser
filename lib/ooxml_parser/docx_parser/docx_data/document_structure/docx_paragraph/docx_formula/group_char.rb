# Group Char Data
module OoxmlParser
  class GroupChar
    attr_accessor :symbol, :position, :vertical_align, :element

    def initialize(symbol = nil, position = nil, vertical_align = nil, element = nil)
      @symbol = symbol
      @position = position
      @vertical_align = vertical_align
      @element = element
    end
  end
end
