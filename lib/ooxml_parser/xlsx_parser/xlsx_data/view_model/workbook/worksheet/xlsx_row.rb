require_relative 'xlsx_row/xlsx_cell'
module OoxmlParser
  # Single Row of XLSX
  class XlsxRow < OOXMLDocumentObject
    attr_accessor :cells, :height, :style, :hidden

    def initialize(parent: nil)
      @cells = []
      @parent = parent
    end

    # Parse XlsxRow object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [XlsxRow] result of parsing
    def parse(node)
      @height = node.attribute('ht').value if option_enabled?(node, 'customHeight') && node.attribute('ht')
      @hidden = option_enabled?(node, 'hidden')
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'c'
          @cells[Coordinates.parse_coordinates_from_string(node_child.attribute('r').value.to_s).get_column_number.to_i - 1] = XlsxCell.parse(node_child)
        end
      end
      self
    end
  end
end
