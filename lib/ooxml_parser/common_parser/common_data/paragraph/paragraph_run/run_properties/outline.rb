# frozen_string_literal: true

# Class for parsing Run Outline properties
module OoxmlParser
  # Class for parsing `w:ln` tags
  class Outline < OOXMLDocumentObject
    # @return [OoxmlSize] width of outline
    attr_reader :width
    # @return [DocxColorScheme] color of outline
    attr_reader :color_scheme

    def initialize(parent: nil)
      @width = OoxmlSize.new(0)
      @parent = parent
    end

    # Parse Outline object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Outline] result of parsing
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
