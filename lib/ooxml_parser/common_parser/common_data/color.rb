require_relative 'color/color_helper'
require_relative 'color/ooxml_color'
require_relative 'colors/color_alpha_channel'
require_relative 'colors/hsl_color'
require_relative 'colors/scheme_color'
require_relative 'colors/color_properties'
require_relative 'colors/theme_colors'
# @author Pavel.Lobashov
# Class for Color in RGB
module OoxmlParser
  class Color < OOXMLDocumentObject
    include ColorHelper
    # @return [Array] Deprecated Indexed colors
    # List of color duplicated from `OpenXML Sdk IndexedColors` class
    # See https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.indexedcolors.aspx
    COLOR_INDEXED =
      %w[
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
      ].freeze

    # @return [Integer] Value of Red Part
    attr_accessor :red
    # @return [Integer] Value of Green Part
    attr_accessor :green
    # @return [Integer] Value of Blue Part
    attr_accessor :blue
    # @return [String] Value of Color Style
    attr_accessor :style
    alias set_style style=
    # @return [String] color scheme of color
    attr_accessor :scheme

    # @return [Integer] Value of alpha-channel
    attr_accessor :alpha_channel
    alias set_alpha_channel alpha_channel=

    attr_accessor :position
    attr_accessor :properties

    # Value of color if non selected
    VALUE_FOR_NONE_COLOR = nil

    def initialize(new_red = VALUE_FOR_NONE_COLOR,
                   new_green = VALUE_FOR_NONE_COLOR,
                   new_blue = VALUE_FOR_NONE_COLOR,
                   parent: nil)
      @red = new_red
      @green = new_green
      @blue = new_blue
      @parent = parent
    end

    def to_s
      if @red == VALUE_FOR_NONE_COLOR && @green == VALUE_FOR_NONE_COLOR && @blue == VALUE_FOR_NONE_COLOR
        'none'
      else
        "RGB (#{@red}, #{@green}, #{@blue})"
      end
    end

    def inspect
      to_s
    end

    def to_hex
      (@red.to_s(16).rjust(2, '0') + @green.to_s(16).rjust(2, '0') + @blue.to_s(16).rjust(2, '0')).upcase
    end

    alias to_int16 to_hex

    def none?
      (@red == VALUE_FOR_NONE_COLOR) && (@green == VALUE_FOR_NONE_COLOR) && (@blue == VALUE_FOR_NONE_COLOR)
    end

    def any?
      !none?
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
        elsif (@red == other.red) && (@green == other.green) && (@blue == other.blue)
          true
        elsif (none? && other.white?) || (white? && other.none?)
          true
        else
          false
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
      color_to_check = color_to_check.pattern_fill.foreground_folor if color_to_check.is_a?(Fill)
      color_to_check = color_to_check.color.converted_color if color_to_check.is_a?(PresentationFill)
      color_to_check = Color.parse(color_to_check) if color_to_check.is_a?(String)
      color_to_check = Color.parse(color_to_check.to_s) if color_to_check.is_a?(Symbol)
      color_to_check = Color.parse(color_to_check.value) if color_to_check.is_a?(DocxColor)
      return true if none? && color_to_check.nil?
      return false if none? && color_to_check.any?
      return false if !none? && color_to_check.none?
      return true if self == color_to_check
      red = color_to_check.red
      green = color_to_check.green
      blue = color_to_check.blue
      (self.red - red).abs < delta && (self.green - green).abs < delta && (self.blue - blue).abs < delta ? true : false
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

    class << self
      def generate_random_color
        Color.new(rand(256), rand(256), rand(256))
      end

      alias random generate_random_color

      # Read array of color from the AllTestData's constant
      # @param [Array] const_array_name - array of the string
      # @return [Array, Color] array of color
      def array_from_const(const_array_name)
        const_array_name.map { |current_color| Color.parse_string(current_color) }
      end

      def get_rgb_by_color_index(index)
        color_by_index = COLOR_INDEXED[index]
        return :unknown if color_by_index.nil?
        color_by_index == 'n/a' ? Color.new : Color.new.parse_hex_string(color_by_index)
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

      alias parse parse_string

      def to_color(something)
        case something
        when SchemeColor
          something.converted_color
        when DocxColorScheme
          something.color
        when Fill
          something.to_color
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

      def parse_color_tag(node)
        return if node.nil?
        node.attributes.each do |key, value|
          case key
          when 'val'
            return Color.new.parse_hex_string(value.value.to_s)
          when 'theme'
            return ThemeColors.parse_color_theme(node)
          when 'rgb'
            return Color.new.parse_hex_string(value.value.to_s)
          when 'indexed'
            return Color.get_rgb_by_color_index(value.value.to_i)
          end
        end
      end

      def parse_scheme_color(scheme_color_node)
        color = ThemeColors.list[scheme_color_node.attribute('val').value.to_sym]
        hls = HSLColor.rgb_to_hsl(color)
        scheme_name = nil
        scheme_color_node.xpath('*').each do |scheme_color_node_child|
          case scheme_color_node_child.name
          when 'lumMod'
            hls.l = hls.l * (scheme_color_node_child.attribute('val').value.to_f / 100_000.0)
          when 'lumOff'
            hls.l = hls.l + (scheme_color_node_child.attribute('val').value.to_f / 100_000.0)
          end
        end
        scheme_color_node.attributes.each do |key, value|
          case key
          when 'val'
            scheme_name = value.to_s
          end
        end
        color = hls.to_rgb
        color.alpha_channel = ColorAlphaChannel.parse(scheme_color_node)
        color.scheme = scheme_name
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
            color = Color.new.parse_hex_string(color_model_node.attribute('val').value)
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
          color = Color.new.parse_hex_string(color_node.attribute('val').value)
          color.properties = ColorProperties.new(parent: color).parse(color_node)
          color
        when 'schemeClr'
          color = SchemeColor.new
          color.value = ThemeColors.list[color_node.attribute('val').value.to_sym]
          color.properties = ColorProperties.new(parent: color).parse(color_node)
          color.converted_color = parse_scheme_color(color_node)
          color.value.calculate_with_tint!(1.0 - color.properties.tint) if color.properties.tint
          color
        end
      end
    end
  end
end
