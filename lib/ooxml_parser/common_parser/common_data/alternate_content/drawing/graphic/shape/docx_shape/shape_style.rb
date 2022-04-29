# frozen_string_literal: true

require_relative 'shape_style/style_matrix_reference'
require_relative 'shape_style/font_reference'
module OoxmlParser
  # Class for parsing `wps:style` tags
  class ShapeStyle < OOXMLDocumentObject
    # @return [FillReference] effect reference
    attr_reader :effect_reference
    # @return [FillReference] fill reference
    attr_reader :fill_reference
    # @return [FontReference] font reference
    attr_reader :font_reference
    # @return [LineReference] line reference
    attr_reader :line_reference

    # Parse ShapeStyle object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ShapeStyle] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'effectRef'
          @effect_reference = StyleMatrixReference.new(parent: self).parse(node_child)
        when 'fillRef'
          @fill_reference = StyleMatrixReference.new(parent: self).parse(node_child)
        when 'fontRef'
          @font_reference = FontReference.new(parent: self).parse(node_child)
        when 'lnRef'
          @line_reference = StyleMatrixReference.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
