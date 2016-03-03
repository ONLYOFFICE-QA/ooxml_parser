require_relative 'table_cell_line'
# Border Data
module OoxmlParser
  class Border
    attr_accessor :style, :color

    def initialize(style = nil, color = nil)
      @style = style
      @color = color
    end

    def ==(other)
      return true if nil? && other.nil?
      return false if nil? && !other.nil?
      return false if !nil? && other.nil?
      return false if style != other.style
      return false if color != other.color
      true
    end

    def to_s
      "Border style: #{style}, color: #{color}"
    end
  end
end
