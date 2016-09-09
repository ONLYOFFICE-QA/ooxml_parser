module OoxmlParser
  # Class for parsing offset of drawing `xdr:from`, `xdr:to`
  class XlsxDrawingPositionParameters < OOXMLDocumentObject
    attr_accessor :column, :column_offset, :row, :row_offset

    # Parse XlsxDrawingPositionParameters object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [XlsxDrawingPositionParameters] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'col'
          @column = Coordinates.get_column_name(node_child.text.to_i + 1)
        when 'colOff'
          @column_offset = OoxmlSize.new(node_child.text.to_f, :emu)
        when 'row'
          @row = node_child.text.to_i + 1
        when 'rowOff'
          @row_offset = OoxmlSize.new(node_child.text.to_f, :emu)
        end
      end
      self
    end
  end
end
