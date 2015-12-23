# Delimiter Data
module OoxmlParser
  class Delimiter
    attr_accessor :begin_character, :value, :end_character

    def initialize(begin_character = '(', value = nil, end_character = ')')
      @begin_character = begin_character
      @value = value
      @end_character = end_character
    end
  end
end
