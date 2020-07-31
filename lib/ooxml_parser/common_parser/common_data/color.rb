# frozen_string_literal: true

require_relative 'color/color_helper'
require_relative 'color/ooxml_color'
require_relative 'colors/color_alpha_channel'
require_relative 'colors/hsl_color'
require_relative 'colors/scheme_color'
require_relative 'colors/color_properties'
require_relative 'colors/theme_colors'
# @author Pavel.Lobashov
module OoxmlParser
  # Class for Color in RGB
  class Color < OOXMLDocumentObject
    include ColorHelper
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

    # @return [ColorProperties] properties of color
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

    # @return [String] result of convert of object to string
    def to_s
      if @red == VALUE_FOR_NONE_COLOR && @green == VALUE_FOR_NONE_COLOR && @blue == VALUE_FOR_NONE_COLOR
        'none'
      else
        "RGB (#{@red}, #{@green}, #{@blue})"
      end
    end

    # @return [String] inspect of object for debug means
    def inspect
      to_s
    end

    # @return [String] color in hex value
    def to_hex
      (@red.to_s(16).rjust(2, '0') + @green.to_s(16).rjust(2, '0') + @blue.to_s(16).rjust(2, '0')).upcase
    end

    alias to_int16 to_hex

    # @return [True, False] is color default
    def none?
      (@red == VALUE_FOR_NONE_COLOR) && (@green == VALUE_FOR_NONE_COLOR) && (@blue == VALUE_FOR_NONE_COLOR) ||
        (style == :nil)
    end

    # @return [True, False] is color not default
    def any?
      !none?
    end

    # @return [True, False] is color white
    def white?
      (@red == 255) && (@green == 255) && (@blue == 255)
    end

    # Method to copy object
    # @return [Color] copied object
    def copy
      Color.new(@red, @green, @blue)
    end

    # Compare this object to other
    # @param other [Object] any other object
    # @return [True, False] result of comparision
    def ==(other)
      if other.is_a?(Color)
        if (@red == other.red) && (@green == other.green) && (@blue == other.blue)
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
      return true if none? && color_to_check.none?
      return false if none? && color_to_check.any?
      return false if !none? && color_to_check.none?
      return true if self == color_to_check

      red = color_to_check.red
      green = color_to_check.green
      blue = color_to_check.blue
      (self.red - red).abs < delta && (self.green - green).abs < delta && (self.blue - blue).abs < delta ? true : false
    end

    # Apply tint to color
    # @param tint [Integer] tint to apply
    # @return [void]
    def calculate_with_tint!(tint)
      @red += (tint.to_f * (255 - @red)).to_i
      @green += (tint.to_f * (255 - @green)).to_i
      @blue += (tint.to_f * (255 - @blue)).to_i
    end

    # Apply shade to color
    # @param shade [Integer] shade to apply
    # @return [void]
    def calculate_with_shade!(shade)
      @red = (@red * shade.to_f).to_i
      @green = (@green * shade.to_f).to_i
      @blue = (@blue * shade.to_f).to_i
    end

    # Parse color scheme data
    # @param scheme_color_node [Nokogiri::XML:Element] node to parse
    # @return [Color] result of parsing
    def parse_scheme_color(scheme_color_node)
      color = root_object.theme.color_scheme[scheme_color_node.attribute('val').value.to_sym].color
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
      @red = color.red
      @green = color.green
      @blue = color.blue
      @alpha_channel = ColorAlphaChannel.parse(scheme_color_node)
      @scheme = scheme_name
      self
    end

    # Parse color model data
    # @param color_model_parent_node [Nokogiri::XML:Element] node to parse
    # @return [Color] result of parsing
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
        when 'srgbClr'
          color = Color.new.parse_hex_string(color_model_node.attribute('val').value)
          color.alpha_channel = ColorAlphaChannel.parse(color_model_node)
        when 'schemeClr'
          color = Color.new(parent: self).parse_scheme_color(color_model_node)
        end
      end
      return nil unless color

      color.calculate_with_tint!(1.0 - tint) if tint
      @red = color.red
      @green = color.green
      @blue = color.blue
      @alpha_channel = color.alpha_channel
      @scheme = color.scheme
      self
    end

    # Parse color data
    # @param color_node [Nokogiri::XML:Element] node to parse
    # @return [Color] result of parsing
    def parse_color(color_node)
      case color_node.name
      when 'srgbClr'
        color = parse_hex_string(color_node.attribute('val').value)
        color.properties = ColorProperties.new(parent: color).parse(color_node)
        color
      when 'schemeClr'
        color = SchemeColor.new(parent: parent)
        return ValuedChild.new(:string, parent: parent).parse(color_node) unless root_object.theme

        color.value = root_object.theme.color_scheme[color_node.attribute('val').value.to_sym].color
        color.properties = ColorProperties.new(parent: color).parse(color_node)
        color.converted_color = Color.new(parent: self).parse_scheme_color(color_node)
        color.value.calculate_with_tint!(1.0 - color.properties.tint) if color.properties.tint
        color
      end
    end

    class << self
      # @return [Color] random color
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

      # @param index [Integer] index to get
      # @return [Color] color by it's index
      def get_rgb_by_color_index(index)
        color_by_index = color_indexes[index]
        return :unknown if color_by_index.nil?

        color_by_index == 'n/a' ? Color.new : Color.new.parse_hex_string(color_by_index)
      end

      # @return [Array] Deprecated Indexed colors
      #   List of color duplicated from `OpenXML Sdk IndexedColors` class
      #   See https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.indexedcolors.aspx
      def color_indexes
        @color_indexes ||= File.readlines("#{__dir__}/color/color_indexes.list", chomp: true)
      end

      # Parse color from string
      # @param str [String] string to parse
      # @return [Color] result of parsing
      def parse_string(str)
        return str if str.is_a?(Color)
        return Color.new(VALUE_FOR_NONE_COLOR, VALUE_FOR_NONE_COLOR, VALUE_FOR_NONE_COLOR) if str == 'none' || str == '' || str == 'transparent' || str.nil?

        split = if str.include?('RGB (') || str.include?('rgb(')
                  str.gsub(/[(RGBrgb() )]/, '').split(',')
                elsif str.include?('RGB ') || str.include?('rgb')
                  str.gsub(/RGB |rgb/, '').split(', ')
                else
                  raise "Incorrect data for color to parse: '#{str}'"
                end

        Color.new(split[0].to_i, split[1].to_i, split[2].to_i)
      end

      alias parse parse_string

      # Convert other object type to Color
      # @param something [Object] object to convert
      # @return [Color] result of conversion
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
    end
  end
end
