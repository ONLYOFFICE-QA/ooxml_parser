require_relative 'underline'
# @author Pavel.Lobashov
# Class for working with font styles (bold,italic,underlined,strike)
# noinspection RubyClassMethodNamingConvention
module OoxmlParser
  class FontStyle
    # @return [false,true] is bold?
    attr_accessor :bold
    # @return [false,true] is bold?
    attr_accessor :italic
    # @return [Underline] underline type
    attr_accessor :underlined
    # @return [Strike] strike type
    attr_accessor :strike

    # Default constructor
    # @param [true, false] bold is bold?
    # @param [true, false] italic is italic?
    # @param [true, false, String] underlined if not false or nil - default Underline, else none
    # @param [Symbol] strike string with strike type
    # @return [FontStyle] new font style
    def initialize(bold = false, italic = false, underlined = Underline.new(:none), strike = :none)
      @bold = bold
      @italic = italic
      @underlined = underlined == false || underlined.nil? ? Underline.new(:none) : underlined
      @strike = strike
    end

    class << self
      # Generate all possible combination of styles without parameter stoke (8)
      # @return [Array] with all FontStyle
      def generate_all_styles_without_stroke
        styles_array = []
        styles_array << FontStyle.new(true, true, Underline.new(:single), :none)
        styles_array << FontStyle.new(true, true, nil, :none)
        styles_array << FontStyle.new(true, false, Underline.new(:single), :none)
        styles_array << FontStyle.new(true, false, nil, :none)

        styles_array << FontStyle.new(false, true, Underline.new(:single), :none)
        styles_array << FontStyle.new(false, true, nil, :none)
        styles_array << FontStyle.new(false, false, Underline.new(:single), :none)
        styles_array << FontStyle.new(false, false, nil, :none)
      end

      # Generate all possible combination of styles (16)
      # @return [Array] with all FontStyle
      def generate_all_styles
        styles_array = []
        styles_array << generate_all_styles_without_stroke

        styles_array << FontStyle.new(true, true, Underline.new(:single), :single)
        styles_array << FontStyle.new(true, true, nil, :single)
        styles_array << FontStyle.new(true, false, Underline.new(:single), :single)
        styles_array << FontStyle.new(false, true, Underline.new(:single), :single)

        styles_array << FontStyle.new(true, false, nil, :single)
        styles_array << FontStyle.new(false, false, Underline.new(:single), :single)
        styles_array << FontStyle.new(false, true, nil, :single)
        styles_array << FontStyle.new(false, false, nil, :single)
        styles_array.flatten
      end
    end

    # Default == operator
    # @return [true, false] true if two same, false if different
    def ==(other)
      (@bold == other.bold) && (@italic == other.italic) && (@underlined == other.underlined) && (@strike == other.strike)
    end

    # Default + operator
    # @return [FontStyle] the result of overlapping of two styles
    def +(other)
      result = FontStyle.new
      result.bold = other.bold == @bold ? @bold : other.bold
      result.italic = other.italic == @italic ? @italic : other.italic
      result.underlined = other.underlined == @underlined ? @underlined : other.underlined
      result.strike = other.strike == @strike ? @strike : other.strike
      result
    end

    # Default to_s operator
    # @return [String] with text representation
    def to_s
      "Bold: #{@bold}, Italic: #{@italic}, Underlined: #{@underlined}, Strike: #{@strike}"
    end

    def copy
      return dup if underlined.nil?
      other = FontStyle.new
      instance_variables.each do |variable|
        case variable
        when @underlined
          other.instance_variable_set(variable, underlined.copy)
        else
          other.instance_variable_set(variable, instance_variable_get(variable))
        end
      end
      other
    end
  end
end
