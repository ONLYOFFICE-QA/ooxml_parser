# frozen_string_literal: true

module OoxmlParser
  # Docx Shape Line Element
  class DocxShapeLineElement < OOXMLDocumentObject
    attr_accessor :type, :points

    def initialize(points = [], parent: nil)
      @points = points
      super(parent: parent)
    end

    # Parse DocxShapeLineElement
    # @param [Nokogiri::XML:Node] node with DocxShapeLineElement
    # @return [DocxShapeLineElement] result of parsing
    def parse(node)
      case node.name
      when 'moveTo'
        @type = :move
      when 'lnTo'
        @type = :line
      when 'arcTo'
        @type = :arc
      when 'cubicBezTo'
        @type = :cubic_bezier
      when 'close'
        @type = :close
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'pt'
          @points << OOXMLCoordinates.parse(node_child)
        end
      end
      self
    end
  end
end
