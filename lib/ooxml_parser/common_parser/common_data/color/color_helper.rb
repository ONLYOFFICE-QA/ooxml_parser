# frozen_string_literal: true

module OoxmlParser
  # Helper methods for color
  module ColorHelper
    # Parse string in hex
    # @param [String] hex_string with or without alpha-channel
    def parse_hex_string(hex_string)
      return self if %w[auto null].include?(hex_string)

      char_array = hex_string.chars
      case char_array.length
      when 3
        @red = char_array[0].hex
        @green = char_array[1].hex
        @blue = char_array[2].hex
      when 6
        @red = (char_array[0] + char_array[1]).hex
        @green = (char_array[2] + char_array[3]).hex
        @blue = (char_array[4] + char_array[5]).hex
      when 8
        @alpha_channel = (char_array[0] + char_array[1]).hex
        @red = (char_array[2] + char_array[3]).hex
        @green = (char_array[4] + char_array[5]).hex
        @blue = (char_array[6] + char_array[7]).hex
      end
      self
    end

    # Convert color to HSLColor
    # @return [HSLColor]
    def to_hsl
      hls_color = HSLColor.new

      min = [red, green, blue].min.to_f
      max = [red, green, blue].max.to_f

      delta = (max - min).to_f
      hls_color.l = (min + max) / 255.0 / 2.0
      hls_color.alpha_channel = alpha_channel.to_f / 255.0

      unless max == min
        hls_color.s = delta / (255.0 - (255.0 - max - min).abs)

        hls_color.h = if max == red && green >= blue
                        60.0 * (green - blue) / delta
                      elsif max == red && green < blue
                        (60.0 * (green - blue) / delta) + 360.0
                      elsif max == green
                        (60.0 * (blue - red) / delta) + 120.0
                      else
                        (60.0 * (red - green) / delta) + 240.0
                      end
      end

      hls_color
    end
  end
end
