# frozen_string_literal: true

module OoxmlParser
  # HSL is one of the most common cylindrical-coordinate representations of points in an RGB color model.
  # HSL stands for hue, saturation, and lightness, and is often also called HLS.
  class HSLColor
    attr_accessor :h, :s, :l
    # @return [Integer] alpha channel value
    attr_accessor :alpha_channel

    # Hue - The "attribute of a visual sensation according to which an area appears to be similar to one of
    # the perceived colors: red, yellow, green, and blue, or to a combination of two of them".
    # Saturation - The "colorfulness of a stimulus relative to its own brightness".
    # Lightness - The "brightness relative to the brightness of a similarly illuminated white".
    def initialize(hue = 0, saturation = 0, lightness = 0, alpha_channel = nil)
      @alpha_channel = alpha_channel
      @h = hue
      @s = saturation
      @l = lightness
    end

    # Chroma - The "colorfulness relative to the brightness of a similarly illuminated white".
    # @return [Color] result
    def to_rgb
      chroma = (1 - ((2 * @l) - 1).abs) * @s
      x = chroma * (1 - (((@h / 60.0) % 2.0) - 1).abs)
      m = @l - (chroma / 2.0)

      rgb = if @h.zero?
              Color.new(0, 0, 0)
            elsif @h.positive? && @h < 60
              Color.new(chroma, x, 0)
            elsif @h > 60 && @h < 120
              Color.new(x, chroma, 0)
            elsif @h > 120 && @h < 180
              Color.new(0, chroma, x)
            elsif @h > 180 && @h < 240
              Color.new(0, x, chroma)
            elsif @h > 240 && @h < 300
              Color.new(x, 0, chroma)
            else
              Color.new(chroma, 0, x)
            end
      Color.new(((rgb.red + m) * 255.0).round, ((rgb.green + m) * 255.0).round, ((rgb.blue + m) * 255.0).round)
    end

    # Get lum value of color
    # @param tint [Integer] tint to apply
    # @param lum [Integer] lum without tint
    # @return [Integer] result
    def calculate_lum_value(tint, lum)
      if tint.nil?
        lum
      else
        tint.negative? ? lum * (1.0 + tint) : (lum * (1.0 - tint)) + (255 - (255 * (1.0 - tint)))
      end
    end

    # Convert to rgb with applied tint
    # @param tint [Integer] tint to apply
    # @return [Color] result of covert
    def calculate_rgb_with_tint(tint)
      self.l = calculate_lum_value(tint, @l * 255.0) / 255.0
      to_rgb
    end
  end
end
