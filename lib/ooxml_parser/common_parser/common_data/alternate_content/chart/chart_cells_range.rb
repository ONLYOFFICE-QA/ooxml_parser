# frozen_string_literal: true

module OoxmlParser
  # Cell Range of Chart
  class ChartCellsRange < OOXMLDocumentObject
    attr_accessor :list, :points

    def initialize(parent: nil)
      @list = ''
      @points = []
      super
    end

    # Parse ChartCellsRange object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ChartCellsRange] result of parsing
    def parse(node)
      @list = node.xpath('c:f')[0].text.split('!').first
      coordinates = Coordinates.parser_coordinates_range(node.xpath('c:f')[0].text) # .split('!')[1].gsub('$', ''))
      return self unless coordinates

      node.xpath('c:numCache/c:pt').each_with_index do |point_node, index|
        point = ChartPoint.new(coordinates[index])
        point.value = point_node.xpath('c:v').first.text.to_f unless point_node.xpath('c:v').first.nil?
        @points << point
      end
      self
    end
  end
end
