# frozen_string_literal: true

require_relative 'docx_shape_line_path/docx_shape_line_element'
module OoxmlParser
  # Docx Shape Line Path
  class DocxShapeLinePath < OOXMLDocumentObject
    attr_accessor :width, :height, :fill, :stroke, :elements

    def initialize(elements = [], parent: nil)
      @elements = elements
      super(parent: parent)
    end

    # Parse DocxShapeLinePath object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocxShapeLinePath] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'w'
          @width = value.value.to_f
        when 'h'
          @height = value.value.to_f
        when 'stroke'
          @stroke = value.value.to_f
        end
      end
      node.xpath('*').each do |node_child|
        @elements << DocxShapeLineElement.new(parent: self).parse(node_child)
      end
      self
    end
  end
end
