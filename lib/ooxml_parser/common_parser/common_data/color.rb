require_relative 'colors/color_alpha_channel'
require_relative 'colors/hsl_color'
require_relative 'colors/scheme_color'
require_relative 'colors/color_properties'
require_relative 'colors/theme_colors'
# @author Pavel.Lobashov
# Class for Color in RGB
module OoxmlParser
  class Color
    # @return [Array] Deprecated Indexed colors
    # List of color duplicated from `OpenXML Sdk IndexedColors` class
    # See https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.indexedcolors.aspx
    COLOR_INDEXED =
        %w(
          000000
          FFFFFF
          FF0000
          00FF00
          0000FF
          FFFF00
          FF00FF
          00FFFF
          000000
          FFFFFF
          FF0000
          00FF00
          0000FF
          FFFF00
          FF00FF
          00FFFF
          800000
          008000
          000080
          808000
          800080
          008080
          C0C0C0
          808080
          9999FF
          993366
          FFFFCC
          CCFFFF
          660066
          FF8080
          0066CC
          CCCCFF
          000080
          FF00FF
          FFFF00
          00FFFF
          800080
          800000
          008080
          0000FF
          00CCFF
          CCFFFF
          CCFFCC
          FFFF99
          99CCFF
          FF99CC
          CC99FF
          FFCC99
          3366FF
          33CCCC
          99CC00
          FFCC00
          FF9900
          FF6600
          666699
          969696
          003366
          339966
          003300
          333300
          993300
          993366
          333399
          333333
          n/a
          n/a
        )

    # @return [Integer] Value of Red Part
    attr_accessor :red
    # @return [Integer] Value of Green Part
    attr_accessor :green
    # @return [Integer] Value of Blue Part
    attr_accessor :blue
    # @return [String] Value of Color Style
    attr_accessor :style
    alias_method :set_style, :style=

    # @return [Integer] Value of alpha-channel
    attr_accessor :alpha_channel
    alias_method :set_alpha_channel, :alpha_channel=

    attr_accessor :position
    attr_accessor :properties

    # Value of color if non selected
    VALUE_FOR_NONE_COLOR = nil

    def initialize(new_red = VALUE_FOR_NONE_COLOR, new_green = VALUE_FOR_NONE_COLOR, new_blue = VALUE_FOR_NONE_COLOR)
      @red = new_red
      @green = new_green
      @blue = new_blue
    end

    def to_s
      if @red == VALUE_FOR_NONE_COLOR && @green == VALUE_FOR_NONE_COLOR && @blue == VALUE_FOR_NONE_COLOR
        'none'
      else
        "RGB (#{@red}, #{@green}, #{@blue})"
      end
    end

    def to_hex
      (@red.to_s(16).rjust(2, '0') + @green.to_s(16).rjust(2, '0') + @blue.to_s(16).rjust(2, '0')).upcase
    end

    alias_method :to_int16, :to_hex

    def none?
      (@red == VALUE_FOR_NONE_COLOR) && (@green == VALUE_FOR_NONE_COLOR) && (@blue == VALUE_FOR_NONE_COLOR)
    end

    def white?
      (@red == 255) && (@green == 255) && (@blue == 255)
    end

    def copy
      Color.new(@red, @green, @blue)
    end

    def ==(other)
      if other.is_a?(Color)
        if nil?
          false
        else
          if (@red == other.red) && (@green == other.green) && (@blue == other.blue)
            true
          elsif (none? && other.white?) || (white? && other.none?)
            true
          else
            false
          end
        end
      else
        false
      end
    end

    # To compare color, which look alike
    # @param [String or Color] color_to_check color to compare
    # @param [Integer] delta max delta for each of specters
    def looks_like?(color_to_check = nil, delta = 8)
      color_to_check = color_to_check.converted_color if color_to_check.is_a?(SchemeColor)
      color_to_check = color_to_check.color if color_to_check.is_a?(ForegroundColor)
      color_to_check = color_to_check.color.converted_color if color_to_check.is_a?(PresentationFill)
      color_to_check = Color.parse(color_to_check) if color_to_check.is_a?(String)
      color_to_check = Color.parse(color_to_check.to_s) if color_to_check.is_a?(Symbol)
      color_to_check = Color.parse(color_to_check.value) if color_to_check.is_a?(DocxColor)
      return false if none? && !color_to_check.none?
      return false if !none? && color_to_check.none?
      return true if self == color_to_check
      red = color_to_check.red
      green = color_to_check.green
      blue = color_to_check.blue
      ((self.red - red).abs < delta && (self.green - green).abs < delta && (self.blue - blue).abs < delta) ? true : false
    end

    def calculate_with_tint!(tint)
      @red += (tint.to_f * (255 - @red)).to_i
      @green += (tint.to_f * (255 - @green)).to_i
      @blue += (tint.to_f * (255 - @blue)).to_i
    end

    def calculate_with_shade!(shade)
      @red = (@red * shade.to_f).to_i
      @green = (@green * shade.to_f).to_i
      @blue = (@blue * shade.to_f).to_i
    end

    def shade(shade_value)
      red = Color.scrgb_to_srgb(Color.shade_for_component(Color.srgb_to_scrgb(@red), shade_value))
      green = Color.scrgb_to_srgb(Color.shade_for_component(Color.srgb_to_scrgb(@green), shade_value))
      blue = Color.scrgb_to_srgb(Color.shade_for_component(Color.srgb_to_scrgb(@blue), shade_value))
      Color.new(red, green, blue)
    end

    def tint(shade_value)
      red = Color.scrgb_to_srgb(Color.tint_for_component(Color.srgb_to_scrgb(@red), shade_value))
      green = Color.scrgb_to_srgb(Color.tint_for_component(Color.srgb_to_scrgb(@green), shade_value))
      blue = Color.scrgb_to_srgb(Color.tint_for_component(Color.srgb_to_scrgb(@blue), shade_value))
      Color.new(red, green, blue)
    end

    # code from js to ruby
    # https://github.com/ONLYOFFICE/ONLYOFFICE-OnlineEditors/blob/master/OfficeWeb/sdk/Excel/model/DrawingObjects/Format/Format.js
    def to_hsl
      max_hls = 255.0
      cd13 = 1.0 / 3.0
      cd23 = 2.0 / 3.0
      hls_color = HSLColor.new
      min = [@red, @green, @blue].min.to_f
      max = [@red, @green, @blue].max.to_f
      delta = (max - min).to_f
      dmax = (max + min) / 255.0
      ddelta = delta / 255.0
      hls_color.l = dmax / 2.0

      if delta != 0
        if hls_color.l < 0.5
          hls_color.s = ddelta / dmax
        else
          hls_color.s = ddelta / (2.0 - dmax)
        end
        ddelta *= 1_530.0
        d_r = (max - @red) / ddelta
        d_g = (max - @green) / ddelta
        d_b = (max - @blue) / ddelta

        if @red == max
          hls_color.h = d_b - d_r
        elsif @green == max
          hls_color.h = cd13 + d_r - d_b
        elsif @blue == max
          hls_color.h = cd23 + d_g - d_r
        end

        hls_color.h += 1 if hls_color.h < 0
        hls_color.h -= 1 if hls_color.h < 1
      end
      hls_color.h = ((hls_color.h * max_hls).round >> 0) & 0xFF
      hls_color.h = 0 if hls_color.h < 0
      hls_color.h = 255 if hls_color.h > 255

      hls_color.l = ((hls_color.l * max_hls).round >> 0) & 0xFF
      hls_color.l = 0 if hls_color.l < 0
      hls_color.l = 255 if hls_color.l > 255

      hls_color.s = ((hls_color.s * max_hls).round >> 0) & 0xFF
      hls_color.s = 0 if hls_color.s < 0
      hls_color.s = 255 if hls_color.s > 255

      hls_color
    end

    # redefine code from our to_hsl realisation to standartised way for services like http://serennu.com/colour/hsltorgb.php
    def to_hsl_standardised
      hsl = to_hsl
      hsl.h = hsl.h * 360 / 256
      hsl.l = (hsl.l * 100.0 / 255.0).round
      hsl.s = (hsl.s * 100.0 / 255.0).round
      hsl
    end

    class << self
      def generate_random_color
        Color.new(rand(256), rand(256), rand(256))
      end

      alias_method :random, :generate_random_color

      def from_int16(color)
        return nil unless color
        return Color.new(nil, nil, nil) if color == 'auto' || color == 'null'

        char_array = color.split(//)
        alpha_channel, red, green, blue = nil
        if char_array.length == 6
          red = (char_array[0] + char_array[1]).hex
          green = (char_array[2] + char_array[3]).hex
          blue = (char_array[4] + char_array[5]).hex
        elsif char_array.length == 8
          alpha_channel = (char_array[0] + char_array[1]).hex
          red = (char_array[2] + char_array[3]).hex
          green = (char_array[4] + char_array[5]).hex
          blue = (char_array[6] + char_array[7]).hex
        end
        font_color = Color.new(red, green, blue)
        font_color.set_alpha_channel(alpha_channel) if alpha_channel
        font_color
      end

      def parse_color_hash(hash)
        Color.new(hash['red'], hash['green'], hash['blue'])
      end

      # Read array of color from the AllTestData's constant
      # @param [Array] const_array_name - array of the string
      # @return [Array, Color] array of color
      def array_from_const(const_array_name)
        const_array_name.map { |current_color| Color.parse_string(current_color) }
      end

      def get_rgb_by_color_index(index)
        color_by_index = COLOR_INDEXED[index]
        color_by_index == 'n/a' ? Color.new : Color.from_int16(color_by_index)
      end

      def parse_string(str)
        return str if str.is_a?(Color)
        return Color.new(VALUE_FOR_NONE_COLOR, VALUE_FOR_NONE_COLOR, VALUE_FOR_NONE_COLOR) if str == 'none' || str == '' || str == 'transparent' || str.nil?
        split = if str.include?('RGB (') || str.include?('rgb(')
                  str.gsub(/[(RGBrgb\(\) )]/, '').split(',')
                elsif str.include?('RGB ') || str.include?('rgb')
                  str.gsub(/RGB |rgb/, '').split(', ')
                else
                  raise "Incorrect data for color to parse: '#{str}'"
                end

        Color.new(split[0].to_i, split[1].to_i, split[2].to_i)
      end

      alias_method :parse, :parse_string

      def parse_int16_string(color)
        return nil if color.nil?

        char_array = color.split(//)
        red = (char_array[0] + char_array[1]).hex
        green = (char_array[2] + char_array[3]).hex
        blue = (char_array[4] + char_array[5]).hex
        Color.new(red, green, blue)
      end

      def srgb_to_scrgb(color)
        lineal_value = color.to_f / 255.0
        if lineal_value < 0
          result_color = 0
        elsif lineal_value <= 0.04045
          result_color = lineal_value / 12.92
        elsif lineal_value <= 1
          result_color = ((lineal_value + 0.055) / 1.055)**2.4
        else
          result_color = 1
        end
        result_color
      end

      def scrgb_to_srgb(color)
        if color < 0
          result_color = 0
        elsif color <= 0.0031308
          result_color = color * 12.92
        elsif color < 1
          result_color = 1.055 * (color**(1.0 / 2.4)) - 0.055
        else
          result_color = 1
        end
        (result_color * 255.0).to_i
      end

      def shade_for_component(color_component, shade_value)
        if color_component * shade_value < 0
          0
        else
          if color_component * shade_value > 1
            1
          else
            color_component * shade_value
          end
        end
      end

      def tint_for_component(color_component, tint)
        if tint > 0
          color_component * (1 - tint) + tint
        else
          color_component * (1 + tint)
        end
      end

      def to_color(something)
        case something
        when SchemeColor
          something.converted_color
        when ForegroundColor, DocxColorScheme
          something.color
        when PresentationFill
          if something.color.respond_to? :converted_color
            something.color.converted_color
          else
            something.color
          end
        when String
          Color.parse(something)
        when Symbol
          Color.parse(something.to_s)
        when DocxColor
          Color.parse(something.value)
        else
          something
        end
      end

      def parse_color_tag(color_tag)
        unless color_tag.nil?
          if color_tag.attribute('indexed').nil? && !color_tag.attribute('rgb').nil?
            if color_tag.attribute('rgb').value.is_a?(String)
              return Color.from_int16(color_tag.attribute('rgb').value.to_s)
            end
            return Color.from_int16(color_tag.attribute('rgb').value.value.to_s)
          elsif color_tag.attribute('rgb').nil? && !color_tag.attribute('indexed').nil?
            if color_tag.attribute('indexed').value.is_a?(String)
              return Color.get_rgb_by_color_index(color_tag.attribute('indexed').value.to_i)
            end
            return Color.get_rgb_by_color_index(color_tag.attribute('indexed').value.value.to_i)
          elsif !color_tag.attribute('theme').nil?
            return ThemeColors.parse_color_theme(color_tag)
          end
        end
        nil
      end

      def parse_scheme_color(scheme_color_node)
        color = ThemeColors.list[scheme_color_node.attribute('val').value.to_sym]
        hls = HSLColor.rgb_to_hsl(color)
        scheme_color_node.xpath('*').each do |scheme_color_node_child|
          case scheme_color_node_child.name
          when 'lumMod'
            hls.l = hls.l * (scheme_color_node_child.attribute('val').value.to_f / 100_000.0)
          when 'lumOff'
            hls.l = hls.l + (scheme_color_node_child.attribute('val').value.to_f / 100_000.0)
          end
        end
        color = hls.to_rgb
        color.alpha_channel = ColorAlphaChannel.parse(scheme_color_node)
        color
      end

      def parse_color_model(color_model_parent_node)
        color = nil
        tint = nil
        color_model_parent_node.xpath('*').each do |color_model_node|
          color_model_node.xpath('*').each do |color_mode_node_child|
            case color_mode_node_child.name
            when 'tint'
              tint = color_mode_node_child.attribute('val').value.to_f / 100_000.0
            end
          end
          case color_model_node.name
          when 'scrgbClr'
            color = Color.new(color_model_node.attribute('r').value.to_i, color_model_node.attribute('g').value.to_i, color_model_node.attribute('b').value.to_i)
            color.alpha_channel = ColorAlphaChannel.parse(color_model_node)
          when 'srgbClr'
            color = Color.from_int16(color_model_node.attribute('val').value)
            color.alpha_channel = ColorAlphaChannel.parse(color_model_node)
          when 'schemeClr'
            color = parse_scheme_color(color_model_node)
          end
        end
        return nil unless color
        color.calculate_with_tint!(1.0 - tint) if tint
        color
      end

      def parse_color(color_node)
        case color_node.name
        when 'srgbClr'
          color = Color.from_int16(color_node.attribute('val').value)
          color.properties = ColorProperties.parse(color_node)
          color
        when 'schemeClr'
          color = SchemeColor.new
          color.value = ThemeColors.list[color_node.attribute('val').value.to_sym]
          color.properties = ColorProperties.parse(color_node)
          color.converted_color = HSLColor.calculate_with_luminance color.value,
                                                                    color.properties.luminance_modulation,
                                                                    color.properties.luminance_offset
          color
        end
      end
    end
  end
end
