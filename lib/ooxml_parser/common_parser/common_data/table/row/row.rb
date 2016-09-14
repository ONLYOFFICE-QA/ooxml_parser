require_relative 'cell/cell'
require_relative 'row/table_row_properties'
module OoxmlParser
  class TableRow < OOXMLDocumentObject
    attr_accessor :height, :cells, :table_row_properties

    def initialize(cells = [], parent: nil)
      @cells = cells
      @parent = parent
    end

    alias table_row_height height

    # Parse TableRow object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TableRow] result of parsing
    def parse(node)
      @height = (node.attribute('h').value.to_f / 360_000.0).round(2) unless node.attribute('h').nil?
      Presentation.current_font_style = FontStyle.new(true) # TODO: Add correct parsing of TableStyle.xml file and use it
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'trPr'
          @table_row_properties = TableRowProperties.new(parent: self).parse(node_child)
        when 'tc'
          @cells << TableCell.parse(node_child, parent: self)
        end
      end
      Presentation.current_font_style = FontStyle.new
      self
    end
  end
end
