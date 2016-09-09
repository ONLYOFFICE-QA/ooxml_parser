require_relative 'xlsx_drawing/xlsx_drawing_position_parameters'
require_relative 'xlsx_drawing/xlsx_shape_grouping'
module OoxmlParser
  class XlsxDrawing < OOXMLDocumentObject
    attr_accessor :picture, :shape, :grouping
    # @return [XlsxDrawingPositionParameters] position from
    attr_accessor :from
    # @return [XlsxDrawingPositionParameters] position to
    attr_accessor :to

    # Parse XlsxDrawing object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [XlsxDrawing] result of parsing
    def parse(node)
      node.xpath('*').each do |child_node|
        case child_node.name
        when 'from'
          @from = XlsxDrawingPositionParameters.new(parent: self).parse(child_node)
        when 'to'
          @to = XlsxDrawingPositionParameters.new(parent: self).parse(child_node)
        when 'sp'
          @shape = DocxShape.parse(child_node)
        when 'grpSp'
          @grouping = XlsxShapeGrouping.parse(child_node)
        when 'pic'
          @picture = DocxPicture.parse(child_node)
        when 'graphicFrame'
          @picture = DocxPicture.new
          graphic_data_node = child_node.xpath('a:graphic/a:graphicData', 'xmlns:a' => 'http://schemas.openxmlformats.org/drawingml/2006/main')
          graphic_data_node.xpath('*').each do |graphic_data_node_child|
            case graphic_data_node_child.name
            when 'chart'
              path_to_chart_xml = OOXMLDocumentObject.get_link_from_rels(graphic_data_node_child.attribute('id').value)
              OOXMLDocumentObject.add_to_xmls_stack(path_to_chart_xml)
              @picture.chart = Chart.parse
              OOXMLDocumentObject.xmls_stack.pop
            end
          end
        end
      end
      self
    end
  end
end
