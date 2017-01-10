module OoxmlParser
  class XlsxShapeGrouping
    attr_accessor :properties, :shapes, :pictures, :grouping

    def initialize(shapes = [], pictures = [])
      @shapes = shapes
      @pictures = pictures
    end

    def self.parse(grouping_node)
      grouping = XlsxShapeGrouping.new
      grouping_node.xpath('*').each do |grouping_node_child|
        case grouping_node_child.name
        when 'grpSpPr'
          grouping.properties = DocxShapeProperties.new(parent: grouping).parse(grouping_node_child)
        when 'grpSp'
          grouping.grouping = parse(grouping_node_child)
        when 'sp'
          grouping.shapes << DocxShape.new(parent: grouping).parse(grouping_node_child)
        when 'pic'
          grouping.pictures << DocxPicture.new(parent: grouping).parse(grouping_node_child)
        when 'graphicFrame'
          picture = DocxPicture.new
          graphic_data_node = grouping_node_child.xpath('a:graphic/a:graphicData', 'xmlns:a' => 'http://schemas.openxmlformats.org/drawingml/2006/main')
          graphic_data_node.xpath('*').each do |graphic_data_node_child|
            case graphic_data_node_child.name
            when 'chart'
              picture.chart = Chart.parse
            end
          end
          grouping.pictures << picture
        end
      end
      grouping
    end
  end
end
