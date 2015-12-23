# Table Cell Margin data
module OoxmlParser
  class TableCellMargin
    attr_accessor :left, :right, :top, :bottom

    def initialize(left = nil, right = nil, top = nil, bottom = nil)
      @left = left
      @top = top
      @right = right
      @bottom = bottom
    end

    def ==(other)
      @left == other.left &&
        @top == other.top &&
        @right == other.right &&
        @bottom == other.bottom
    end
  end
end
