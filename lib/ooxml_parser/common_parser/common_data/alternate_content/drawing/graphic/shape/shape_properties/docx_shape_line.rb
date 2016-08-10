require_relative 'color/docx_color_scheme'
require_relative 'line/line_end'
# Docx Shape Line
module OoxmlParser
  class DocxShapeLine
    attr_accessor :width, :color_scheme, :cap, :head_end, :tail_end, :fill

    alias color color_scheme

    def initialize
      @width = 0
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

    def self.parse(shape_line_node)
      shape_line = DocxShapeLine.new
      shape_line.width = (shape_line_node.attribute('w').value.to_f / 12_699.3).round(2) unless shape_line_node.attribute('w').nil?
      unless shape_line_node.attribute('cap').nil?
        case shape_line_node.attribute('cap').value
        when 'rnd'
          shape_line.cap = :round
        when 'sq'
          shape_line.cap = :square
        when 'flat'
          shape_line.cap = :flat
        end
      end
      shape_line_node.xpath('*').each do |shape_line_node_child|
        case shape_line_node_child.name
        when 'solidFill'
          shape_line.color_scheme = DocxColorScheme.parse(shape_line_node_child)
        when 'noFill'
          shape_line.width = 0
        when 'headEnd'
          shape_line.head_end = LineEnd.parse(shape_line_node_child)
        when 'tailEnd'
          shape_line.tail_end = LineEnd.parse(shape_line_node_child)
        end
      end
      shape_line
    end
  end
end
