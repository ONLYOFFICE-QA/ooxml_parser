# Class for working Underline style parameter
module OoxmlParser
  class Underline
    attr_accessor :style, :color

    def initialize(style = :none, color = nil)
      @style = style == 'single' ? :single : style
      @color = color
    end

    def ==(other)
      if other.is_a? Underline
        @style.to_sym == other.style.to_sym && @color == other.color
      elsif other.is_a? Symbol
        @style.to_sym == other
      else
        false
      end
    end

    def to_s
      if @color.nil?
        @style.to_s
      else
        "#{@style} #{@color}"
      end
    end

    def self.parse(attribute_value)
      underline = Underline.new
      case attribute_value
      when 'sng'
        underline.style = :single
      when 'none'
        underline.style = :none
      end
      underline
    end
  end
end
