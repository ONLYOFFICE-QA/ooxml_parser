require_relative 'table_cell_line/line_join'
module OoxmlParser
  # Class for parsing Table Cell Lines
  class TableCellLine < OOXMLDocumentObject
    attr_accessor :fill, :dash, :line_join, :head_end, :tail_end, :align, :width, :cap_type, :compound_line_type

    def initialize(fill = nil, line_join = nil, parent: nil)
      @fill = fill
      @line_join = line_join
      @parent = parent
    end

    # Parse TableCellLine object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TableCellLine] result of parsing
    def parse(node)
      @fill = PresentationFill.new(parent: self).parse(node)
      @line_join = LineJoin.new(parent: self).parse(node)
      node.attributes.each do |key, value|
        case key
        when 'w'
          @width = OoxmlSize.new(value.value.to_f, :emu)
        when 'algn'
          @align = value_to_symbol(value)
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'headEnd'
          @head_end = LineEnd.new(parent: self).parse(node_child)
        when 'tailEnd'
          @tail_end = LineEnd.new(parent: self).parse(node_child)
        when 'ln'
          return TableCellLine.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
