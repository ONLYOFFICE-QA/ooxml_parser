require_relative 'table_cell_line/line_join'
module OoxmlParser
  class TableCellLine < OOXMLDocumentObject
    attr_accessor :fill, :dash, :line_join, :head_end, :tail_end, :align, :width, :cap_type, :compound_line_type

    def initialize(fill = nil, line_join = nil)
      @fill = fill
      @line_join = line_join
    end

    def self.parse(border_node)
      line = TableCellLine.new(PresentationFill.parse(border_node), LineJoin.parse(border_node))
      border_node.xpath('*').each do |border_node_child|
        case border_node_child.name
        when 'prstDash'
        when 'custDash'
        when 'headEnd'
          line.head_end = LineEnd.parse(border_node_child)
        when 'tailEnd'
          line.tail_end = LineEnd.parse(border_node_child)
        when 'ln'
          return TableCellLine.parse(border_node_child)
        end
      end
      border_node.attributes.each do |key, value|
        case key
        when 'w'
          line.width = value.value.to_f / 12_700.0
        when 'algn'
          line.align = Alignment.parse(value)
        end
      end
      line
    end
  end
end
