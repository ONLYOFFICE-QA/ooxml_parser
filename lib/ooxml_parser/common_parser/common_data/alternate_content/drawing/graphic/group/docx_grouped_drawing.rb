require_relative 'docx_group_element'
module OoxmlParser
  # Docx Groping Drawing Data
  class DocxGroupedDrawing < OOXMLDocumentObject
    attr_accessor :elements, :properties

    def initialize(parent: nil)
      @elements = []
      @parent = parent
    end

    # Parse DocxGroupedDrawing object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocxGroupedDrawing] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'grpSpPr'
          @properties = DocxShapeProperties.new(parent: self).parse(node_child)
        when 'pic'
          element = DocxGroupElement.new(:picture, parent: self)
          element.object = DocxPicture.new(parent: element).parse(node_child)
          @elements << element
        when 'wsp'
          element = DocxGroupElement.new(:shape, parent: self)
          element.object = DocxShape.new(parent: element).parse(node_child)
          @elements << element
        when 'grpSp'
          @elements << DocxGroupedDrawing.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
