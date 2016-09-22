require_relative 'table_cell_line'
# Border Data
module OoxmlParser
  class Border
    attr_accessor :style, :color

    def initialize(style = nil, color = nil)
      @style = style
      @color = color
    end
  end
end
