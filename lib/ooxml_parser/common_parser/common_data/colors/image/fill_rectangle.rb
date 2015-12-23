module OoxmlParser
  class FillRectangle
    attr_accessor :left, :top, :right, :bottom

    def self.parse(fill_rectangle_node)
      fill_rectangle = FillRectangle.new
      fill_rectangle_node.attributes.each do |key, value|
        case key
        when 'b'
          fill_rectangle.bottom = value.value.to_i
        when 't'
          fill_rectangle.top = value.value.to_i
        when 'l'
          fill_rectangle.left = value.value.to_i
        when 'r'
          fill_rectangle.right = value.value.to_i
        end
      end
      fill_rectangle
    end
  end
end
