# frozen_string_literal: true

require_relative 'xlsx_drawing/xlsx_drawing_position_parameters'
require_relative 'xlsx_drawing/client_data'
module OoxmlParser
  # Data of spreadsheet drawing
  class XlsxDrawing < OOXMLDocumentObject
    attr_accessor :picture, :shape, :grouping
    # @return [XlsxDrawingPositionParameters] position from
    attr_accessor :from
    # @return [XlsxDrawingPositionParameters] position to
    attr_accessor :to
    # @return [GraphicFrame] graphic frame
    attr_accessor :graphic_frame
    # @return [ClientData] client data
    attr_accessor :client_data

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
          @shape = DocxShape.new(parent: self).parse(child_node)
        when 'grpSp'
          @grouping = ShapesGrouping.new(parent: self).parse(child_node)
        when 'pic'
          @picture = DocxPicture.new(parent: self).parse(child_node)
        when 'graphicFrame'
          @graphic_frame = GraphicFrame.new(parent: self).parse(child_node)
        when 'cxnSp'
          @shape = ConnectionShape.new(parent: self).parse(child_node)
        when 'clientData'
          @client_data = ClientData.new(parent: self).parse(child_node)
        end
      end
      self
    end
  end
end
