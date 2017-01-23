module OoxmlParser
  # Class for parsing `grpSp` tags
  class XlsxShapeGrouping < OOXMLDocumentObject
    attr_accessor :properties, :shapes, :pictures, :grouping

    def initialize(parent: nil)
      @shapes = []
      @pictures = []
      @parent = parent
    end

    # Parse Bookmark object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Bookmark] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'grpSpPr'
          @properties = DocxShapeProperties.new(parent: self).parse(node_child)
        when 'grpSp'
          @grouping = XlsxShapeGrouping.new(parent: self).parse(node_child)
        when 'sp'
          @shapes << DocxShape.new(parent: self).parse(node_child)
        when 'pic'
          @pictures << DocxPicture.new(parent: self).parse(node_child)
        when 'graphicFrame'
          picture = DocxPicture.new
          graphic_data_node = node_child.xpath('a:graphic/a:graphicData', 'xmlns:a' => 'http://schemas.openxmlformats.org/drawingml/2006/main')
          graphic_data_node.xpath('*').each do |graphic_data_node_child|
            case graphic_data_node_child.name
            when 'chart'
              picture.chart = Chart.parse
            end
          end
          @pictures << picture
        end
      end
      self
    end
  end
end
