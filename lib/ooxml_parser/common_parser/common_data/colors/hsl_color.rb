# HSL is one of the most common cylindrical-coordinate representations of points in an RGB color model.
# HSL stands for hue, saturation, and lightness, and is often also called HLS.
module OoxmlParser
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

    def self.rgb_to_hsl(rgb_color)
      hls_color = HSLColor.new
      red = rgb_color.red.to_f # / 255.0
      green = rgb_color.green.to_f # / 255.0
      blue = rgb_color.blue.to_f # / 255.0

      min = [red, green, blue].min.to_f
      max = [red, green, blue].max.to_f

      delta = (max - min).to_f
      hls_color.l = (min + max) / 255.0 / 2.0
      hls_color.alpha_channel = rgb_color.alpha_channel.to_f / 255.0

      unless max == min
        hls_color.s = delta / (255.0 - (255.0 - max - min).abs)

        hls_color.h = if max == red && green >= blue
                        60.0 * (green - blue) / delta
                      elsif max == red && green < blue
                        60.0 * (green - blue) / delta + 360.0
                      elsif max == green
                        60.0 * (blue - red) / delta + 120.0
                      else
                        60.0 * (red - green) / delta + 240.0
                      end
      end

      hls_color
    end

    # Chroma - The "colorfulness relative to the brightness of a similarly illuminated white".
    def to_rgb
      chroma = (1 - (2 * @l - 1).abs) * @s
      x = chroma * (1 - ((@h / 60.0) % 2.0 - 1).abs)
      m = @l - chroma / 2.0

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

    def calculate_lum_value(tint, lum)
      if tint.nil?
        lum
      else
        tint.negative? ? lum * (1.0 + tint) : lum * (1.0 - tint) + (255 - 255 * (1.0 - tint))
      end
    end

    def calculate_rgb_with_tint(tint)
      self.l = calculate_lum_value(tint, @l * 255.0) / 255.0
      to_rgb
    end
  end
end
