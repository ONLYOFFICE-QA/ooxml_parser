# frozen_string_literal: true

# Class for parsing text outline and its properties
module OoxmlParser
  # Class for parsing `w:textOutline`
  class TextOutline < OOXMLDocumentObject
    attr_accessor :width, :color_scheme

    def initialize(parent: nil)
      @width = OoxmlSize.new(0)
      @color_scheme = :none
      @parent = parent
    end

    # Parse TextOutline object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TextOutline] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'w'
          @width = OoxmlSize.new(value.value.to_f, :emu)
        end
      end
      @color_scheme = DocxColorScheme.new(parent: self).parse(node)
      self
    end
  end
end
