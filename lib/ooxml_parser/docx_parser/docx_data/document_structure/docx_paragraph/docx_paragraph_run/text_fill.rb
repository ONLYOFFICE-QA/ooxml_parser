# frozen_string_literal: true

module OoxmlParser
  # Data about `textFill` object
  class TextFill < OOXMLDocumentObject
    attr_accessor :color_scheme

    # Parse TextFill object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TextFill] result of parsing
    def parse(node)
      @color_scheme = DocxColorScheme.new(parent: self).parse(node)
      self
    end
  end
end
