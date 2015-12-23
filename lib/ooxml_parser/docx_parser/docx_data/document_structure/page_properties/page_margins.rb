module OoxmlParser
  class PageMargins
    attr_accessor :top, :bottom, :left, :right, :footer, :gutter, :header

    def initialize
      @top = nil
      @bottom = nil
      @left = nil
      @right = nil
      @header = 1.25
      @footer = 1.25
      @gutter = nil
    end

    def ==(other)
      other.class == PageMargins &&
        @top == other.top &&
        @bottom == other.bottom &&
        @left == other.left &&
        @right == other.right &&
        @footer == other.footer &&
        @header == other.header &&
        @gutter == other.gutter
    end
  end
end
