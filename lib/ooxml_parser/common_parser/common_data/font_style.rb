# frozen_string_literal: true

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

    # Default == operator
    # @return [true, false] true if two same, false if different
    def ==(other)
      (@bold == other.bold) && (@italic == other.italic) && (@underlined == other.underlined) && (@strike == other.strike)
    end

    # Default to_s operator
    # @return [String] with text representation
    def to_s
      "Bold: #{@bold}, Italic: #{@italic}, Underlined: #{@underlined}, Strike: #{@strike}"
    end
  end
end
