# Docx Coordinates
module OoxmlParser
  class OOXMLCoordinates
    attr_accessor :x, :y

    def initialize(x, y)
      @x = if x.is_a?(OoxmlSize)
             x
           else
             OoxmlSize.new(x)
           end
      @y = if y.is_a?(OoxmlSize)
             y
           else
             OoxmlSize.new(y)
           end
    end

    def to_s
      '(' + @x.to_s + '; ' + @y.to_s + ')'
    end

    # Compare two OOXMLCoordinates objects
    # @param other [OOXMLCoordinates] other object
    # @return [True, False] result of comparasion
    def ==(other)
      x == other.x && y == other.y
    end

    # Parse OOXMLCoordinates object
    # @param position_node [Nokogiri::XML:Element] node to parse
    # @param x_attr [String] name of x attribute
    # @param y_attr [String] name of y attribute
    # @param delimiter [Float] delimiter to devise values
    # @return [OOXMLCoordinates] result of parsing
    def self.parse(position_node, x_attr: 'x', y_attr: 'y', unit: :dxa)
      return if position_node.attribute(x_attr).nil? || position_node.attribute(y_attr).nil?
      OOXMLCoordinates.new(OoxmlSize.new(position_node.attribute(x_attr).value.to_f, unit),
                           OoxmlSize.new(position_node.attribute(y_attr).value.to_f, unit))
    end
  end
end
