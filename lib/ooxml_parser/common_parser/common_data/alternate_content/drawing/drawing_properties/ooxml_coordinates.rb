# frozen_string_literal: true

module OoxmlParser
  # Docx Coordinates
  class OOXMLCoordinates < OOXMLDocumentObject
    attr_accessor :x, :y

    def initialize(x_value = nil, y_value = nil, parent: nil)
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
      super(parent: parent)
    end

    # @return [String] result of convert of object to string
    def to_s
      "(#{@x}; #{@y})"
    end

    # Compare two OOXMLCoordinates objects
    # @param other [OOXMLCoordinates] other object
    # @return [True, False] result of comparison
    def ==(other)
      x == other.x && y == other.y
    end

    # Parse OOXMLCoordinates object
    # @param node [Nokogiri::XML:Element] node to parse
    # @param x_attr [String] name of x attribute
    # @param y_attr [String] name of y attribute
    # @param unit [Symbol] unit in which data is stored
    # @return [OOXMLCoordinates] result of parsing
    def parse(node, x_attr: 'x', y_attr: 'y', unit: :dxa)
      node.attributes.each do |key, value|
        case key
        when x_attr
          @x = OoxmlSize.new(value.value.to_f, unit)
        when y_attr
          @y = OoxmlSize.new(value.value.to_f, unit)
        end
      end
      self
    end
  end
end
