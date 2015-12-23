require_relative 'xlsx_drawing_position/xlsx_drawing_position_parameters'
module OoxmlParser
  class XlsxDrawingPosition
    attr_accessor :from, :to

    # Parse XlsxDrawingPosition
    # @param node [Nokogiri::XML::Element] node to parse
    # @return [XlsxDrawingPosition] value of XlsxDrawingPosition
    def self.parse(node)
      position = XlsxDrawingPosition.new
      node.xpath('*').each do |drawing_node_child|
        case drawing_node_child.name
        when 'from'
          position.from = XlsxDrawingPositionParameters.parse(drawing_node_child)
        when 'to'
          position.to = XlsxDrawingPositionParameters.parse(drawing_node_child)
        end
      end
      position
    end
  end
end
