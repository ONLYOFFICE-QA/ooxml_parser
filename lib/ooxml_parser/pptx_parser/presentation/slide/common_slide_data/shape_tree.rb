# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `spTree` tag
  class ShapeTree < OOXMLDocumentObject
    # @return [Array] elements of shape tree
    attr_reader :elements

    def initialize(parent: nil)
      @elements = []
      super
    end

    # Parse ShapeTree object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ShapeTree] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'sp'
          @elements << DocxShape.new(parent: self).parse(node_child).dup
        when 'pic'
          @elements << DocxPicture.new(parent: self).parse(node_child)
        when 'graphicFrame'
          @elements << GraphicFrame.new(parent: self).parse(node_child)
        when 'grpSp'
          @elements << ShapesGrouping.new(parent: self).parse(node_child)
        when 'cxnSp'
          @elements << ConnectionShape.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
