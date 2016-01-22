# Docx Coordinates
module OoxmlParser
  class OOXMLCoordinates
    attr_accessor :x, :y

    def initialize(x, y)
      @x = x
      @y = y
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
    def self.parse(position_node, x_attr: 'x', y_attr: 'y', delimiter: OoxmlParser.configuration.units_delimiter)
      OOXMLCoordinates.new((position_node.attribute(x_attr).value.to_f / delimiter).round(3), (position_node.attribute(y_attr).value.to_f / delimiter).round(3))
    end
  end
end
