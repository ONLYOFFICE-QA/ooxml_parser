# frozen_string_literal: true

module OoxmlParser
  # Docx Drawing Position
  class DocxDrawingPosition < OOXMLDocumentObject
    attr_accessor :relative_from, :offset, :align

    # Parse DocxDrawingPosition object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocxDrawingPosition] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'relativeFrom'
          @relative_from = value_to_symbol(value)
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'posOffset', 'pctPosHOffset', 'pctPosVOffset'
          @offset = OoxmlSize.new(node_child.text.to_f, :emu)
        when 'align'
          @align = node_child.text.to_sym
        end
      end
      self
    end
  end
end
