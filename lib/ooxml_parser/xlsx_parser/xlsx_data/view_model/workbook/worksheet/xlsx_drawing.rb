require_relative 'xlsx_drawing/xlsx_drawing_position'
require_relative 'xlsx_drawing/xlsx_shape_grouping'
module OoxmlParser
  class XlsxDrawing < OOXMLDocumentObject
    attr_accessor :position, :picture, :shape, :grouping

    def initialize(position = XlsxDrawingPosition.new)
      @position = position
    end

    def self.parse(drawing_node)
      drawing = XlsxDrawing.new
      drawing.position = XlsxDrawingPosition.parse(drawing_node)
      drawing_node.xpath('*').each do |drawing_node_child|
        case drawing_node_child.name
        when 'sp'
          drawing.shape = DocxShape.parse(drawing_node_child)
        when 'grpSp'
          drawing.grouping = XlsxShapeGrouping.parse(drawing_node_child)
        when 'pic'
          drawing.picture = DocxPicture.parse(drawing_node_child)
        when 'graphicFrame'
          drawing.picture = DocxPicture.new
          graphic_data_node = drawing_node_child.xpath('a:graphic/a:graphicData', 'xmlns:a' => 'http://schemas.openxmlformats.org/drawingml/2006/main')
          graphic_data_node.xpath('*').each do |graphic_data_node_child|
            case graphic_data_node_child.name
            when 'chart'
              path_to_chart_xml = OOXMLDocumentObject.get_link_from_rels(graphic_data_node_child.attribute('id').value)
              OOXMLDocumentObject.add_to_xmls_stack(path_to_chart_xml)
              drawing.picture.chart = Chart.parse
              OOXMLDocumentObject.xmls_stack.pop
            end
          end
        end
      end
      drawing
    end
  end
end
