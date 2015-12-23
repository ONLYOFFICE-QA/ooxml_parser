# Cell Range of Chart
module OoxmlParser
  class ChartCellsRange
    attr_accessor :list, :points

    def initialize(list = '', points = [])
      @list = list
      @points = points
    end

    def self.parse(num_ref_node)
      chart_range = ChartCellsRange.new
      chart_range.list = num_ref_node.xpath('c:f')[0].text.split('!').first
      coordinates = Coordinates.parser_coordinates_range(num_ref_node.xpath('c:f')[0].text) # .split('!')[1].gsub('$', ''))
      num_ref_node.xpath('c:numCache/c:pt').each_with_index do |point_node, index|
        point = ChartPoint.new(coordinates[index])
        point.value = point_node.xpath('c:v').first.text.to_f unless point_node.xpath('c:v').first.nil?
        chart_range.points << point
      end
      chart_range
    end
  end
end
