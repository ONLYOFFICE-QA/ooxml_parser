# frozen_string_literal: true

module OoxmlParser
  # Docx Pattern Fill Data
  class DocxPatternFill < OOXMLDocumentObject
    attr_accessor :foreground_color, :background_color, :preset

    # Parse DocxPatternFill object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocxPatternFill] result of parsing
    def parse(node)
      @preset = node.attribute('prst').value.to_sym
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'fgClr'
          @foreground_color = Color.new(parent: self).parse_color_model(node_child)
        when 'bgClr'
          @background_color = Color.new(parent: self).parse_color_model(node_child)
        end
      end
      self
    end
  end
end
