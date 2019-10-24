# frozen_string_literal: true

module OoxmlParser
  # Helper methods for color
  module ColorHelper
    # Parse string in hex
    # @param [String] hex_string with or without alpha-channel
    def parse_hex_string(hex_string)
      return self if %w[auto null].include?(hex_string)

      char_array = hex_string.split(//)
      if char_array.length == 3
        @red = char_array[0].hex
        @green = char_array[1].hex
        @blue = char_array[2].hex
      elsif char_array.length == 6
        @red = (char_array[0] + char_array[1]).hex
        @green = (char_array[2] + char_array[3]).hex
        @blue = (char_array[4] + char_array[5]).hex
      elsif char_array.length == 8
        @alpha_channel = (char_array[0] + char_array[1]).hex
        @red = (char_array[2] + char_array[3]).hex
        @green = (char_array[4] + char_array[5]).hex
        @blue = (char_array[6] + char_array[7]).hex
      end
      self
    end
  end
end
