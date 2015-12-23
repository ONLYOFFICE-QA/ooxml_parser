# HSL is one of the most common cylindrical-coordinate representations of points in an RGB color model.
# HSL stands for hue, saturation, and lightness, and is often also called HLS.
module OoxmlParser
  class HSLColor
    attr_accessor :a, :h, :s, :l

    # Hue - The "attribute of a visual sensation according to which an area appears to be similar to one of
    # the perceived colors: red, yellow, green, and blue, or to a combination of two of them".
    # Saturation - The "colorfulness of a stimulus relative to its own brightness".
    # Lightness - The "brightness relative to the brightness of a similarly illuminated white".
    def initialize(hue = 0, saturation = 0, lightness = 0, a = nil)
      @a = a
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
      hls_color.a = rgb_color.alpha_channel.to_f / 255.0

      unless max == min
        hls_color.s = delta / (255.0 - (255.0 - max - min).abs)

        case
        when max == red && green >= blue
          hls_color.h = 60.0 * (green - blue) / delta
        when max == red && green < blue
          hls_color.h = 60.0 * (green - blue) / delta + 360.0
        when max == green
          hls_color.h = 60.0 * (blue - red) / delta + 120.0
        else
          hls_color.h = 60.0 * (red - green) / delta + 240.0
        end
      end

      hls_color
    end

    # Chroma - The "colorfulness relative to the brightness of a similarly illuminated white".
    def to_rgb
      chroma = (1 - (2 * @l - 1).abs) * @s
      x = chroma * (1 - ((@h / 60.0) % 2.0 - 1).abs)
      m = @l - chroma / 2.0

      rgb = case
            when @h == 0
              Color.new(0, 0, 0)
            when @h > 0 && @h < 60
              Color.new(chroma, x, 0)
            when @h > 60 && @h < 120
              Color.new(x, chroma, 0)
            when @h > 120 && @h < 180
              Color.new(0, chroma, x)
            when @h > 180 && @h < 240
              Color.new(0, x, chroma)
            when @h > 240 && @h < 300
              Color.new(x, 0, chroma)
            else
              Color.new(chroma, 0, x)
            end
      Color.new(((rgb.red + m) * 255.0).round, ((rgb.green + m) * 255.0).round, ((rgb.blue + m) * 255.0).round)
    end

    def set_color(t1, t2, t3)
      t3 += 1.0 if t3 < 0
      t3 -= 1.0 if t3 > 1
      case
      when 6.0 * t3 < 1 then
        t2 + (t1 - t2) * 6.0 * t3
      when 2.0 * t3 < 1 then
        t1
      when 3.0 * t3 < 2 then
        t2 + (t1 - t2) * ((2.0 / 3.0) - t3) * 6.0
      else
        t2
      end
    end

    def calculate_lum_value(tint, lum)
      if tint.nil?
        lum
      else
        tint < 0 ? lum * (1.0 + tint) : lum * (1.0 - tint) + (255 - 255 * (1.0 - tint))
      end
    end

    def self.calculate_with_luminance(color, lum_mod, lum_off = nil)
      hls_color = color.is_a?(HSLColor) ? color : HSLColor.rgb_to_hsl(color)
      # for hsl color which have h == 0 need another values of lumOff lumMod - 0.04(+-0.005)
      if lum_mod
        if hls_color.h == 0
          case lum_mod
          when 0.2
            hls_color.l *= 0.169
          when 0.4
            hls_color.l *= 0.369
          when 0.6
            hls_color.l *= 0.55
          when 0.75
            hls_color.l *= 0.705
          when 0.5
            hls_color.l *= 0.469
          else
            hls_color.l *= lum_mod
          end
        else
          hls_color.l *= lum_mod
        end
      end
      current_rgb = hls_color.to_rgb
      return current_rgb if current_rgb == Color.new

      unless lum_off.nil?
        hls_color = HSLColor.rgb_to_hsl(current_rgb)
        # for hsl color which have h == 0 need another values of lumOff - 0.04(+-0.01)
        if hls_color.h == 0
          case lum_off
          when 0.8
            hls_color.l += 0.76
          when 0.6
            hls_color.l += 0.55
          when 0.4
            hls_color.l += 0.36
          else
            hls_color.l += lum_off
          end
        else
          hls_color.l += lum_off
        end
        current_rgb = hls_color.to_rgb
      end
      current_rgb
    end

    def calculate_rgb_with_tint(tint)
      self.l = calculate_lum_value(tint, @l * 255.0) / 255.0
      to_rgb
    end
  end
end
