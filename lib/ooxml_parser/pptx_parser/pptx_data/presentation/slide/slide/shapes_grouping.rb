module OoxmlParser
  # Class for parsing `grpSp`
  class ShapesGrouping < OOXMLDocumentObject
    attr_accessor :elements, :properties

    def initialize(parent: nil)
      @elements = []
      @parent = parent
    end

    # Parse ShapesGrouping object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ShapesGrouping] result of parsing
    def parse(node)
      node.xpath('*').each do |child_node|
        case child_node.name
        when 'cxnSp'
          @elements << ConnectionShape.new(parent: self).parse(child_node)
        when 'grpSpPr'
          @properties = DocxShapeProperties.new(parent: self).parse(child_node)
        when 'pic'
          @elements << DocxPicture.new(parent: self).parse(child_node)
        when 'sp'
          @elements << DocxShape.new(parent: self).parse(child_node).dup
        when 'grpSp'
          @elements << ShapesGrouping.new(parent: self).parse(child_node)
        when 'graphicFrame'
          @elements << GraphicFrame.new(parent: self).parse(child_node)
        when 'wsp'
          @elements << DocxShape.new(parent: self).parse(child_node)
        end
      end
      self
    end
  end
end
