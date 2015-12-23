module OoxmlParser
  class ShapeProperties
    attr_accessor :fill_color, :anchor_x, :anchor_y, :shadow, :margins, :size, :opacity, :position, :z_index, :stroke

    def initialize(position = nil, margins = ShapeMargins.new, size = ShapeSize.new, z_index = nil, stroke = Stroke.new)
      @position = position
      @margins = margins
      @size = size
      @z_index = z_index
      @stroke = stroke
    end
  end
end
