# frozen_string_literal: true

module OoxmlParser
  # Docx Coordinates
  class OOXMLCoordinates
    attr_accessor :x, :y

    def initialize(x_value, y_value)
      @x = if x_value.is_a?(OoxmlSize)
             x_value
           else
             OoxmlSize.new(x_value)
           end
      @y = if y_value.is_a?(OoxmlSize)
             y_value
           else
             OoxmlSize.new(y_value)
           end
    end

    # @return [String] result of convert of object to string
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
    # @param unit [Symbol] unit in which data is stored
    # @return [OOXMLCoordinates] result of parsing
    def self.parse(position_node, x_attr: 'x', y_attr: 'y', unit: :dxa)
      return if position_node.attribute(x_attr).nil? || position_node.attribute(y_attr).nil?

      OOXMLCoordinates.new(OoxmlSize.new(position_node.attribute(x_attr).value.to_f, unit),
                           OoxmlSize.new(position_node.attribute(y_attr).value.to_f, unit))
    end
  end
end
