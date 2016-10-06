require_relative 'color/docx_color_scheme'
require_relative 'line/line_end'
# Docx Shape Line
module OoxmlParser
  class DocxShapeLine < OOXMLDocumentObject
    attr_accessor :width, :color_scheme, :cap, :head_end, :tail_end, :fill

    alias color color_scheme

    def initialize(parent: nil)
      @width = OoxmlSize.new(0)
      @parent = parent
    end

    def stroke_size
      if @color_scheme.nil? || @color_scheme.color == :none
        0
      else
        @width
      end
    end

    def nil?
      stroke_size.zero? && cap.nil?
    end

    # Parse DocxShapeLine object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocxShapeLine] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'w'
          @width = OoxmlSize.new(value.value.to_f, :emu)
        when 'cap'
          @cap = value_to_symbol(value)
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'solidFill'
          @color_scheme = DocxColorScheme.parse(node_child)
        when 'noFill'
          @width = OoxmlSize.new(0)
        when 'headEnd'
          @head_end = LineEnd.new(parent: self).parse(node_child)
        when 'tailEnd'
          @tail_end = LineEnd.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
