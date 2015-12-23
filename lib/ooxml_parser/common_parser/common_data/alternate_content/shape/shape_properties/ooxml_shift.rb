module OoxmlParser
  class OOXMLShift
    attr_accessor :x, :y

    def initialize(x = nil, y = nil)
      @x = x
      @y = y
    end

    def self.parse(shift_node, x_name = 'x', y_name = 'y', divider = 360_000)
      return if shift_node.nil? || shift_node.attribute(x_name).nil? || shift_node.attribute(y_name).nil?
      OOXMLShift.new((shift_node.attribute(x_name).value.to_f / divider.to_f).round(2), (shift_node.attribute(y_name).value.to_f / divider.to_f).round(2))
    end
  end
end
