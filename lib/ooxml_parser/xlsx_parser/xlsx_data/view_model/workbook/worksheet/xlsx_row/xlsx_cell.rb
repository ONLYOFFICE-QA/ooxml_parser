require_relative 'xlsx_cell/cell_style'
# Single Cell of XLSX
module OoxmlParser
  class XlsxCell < OOXMLDocumentObject
    attr_accessor :style, :text, :formula, :character

    # @return [String] text without applying any style modificators, like quote_prefix
    attr_accessor :raw_text

    def initialize(style = nil, text = '', parent: nil)
      @style = style
      @text = text
      @raw_text = ''
      @parent = parent
    end

    # Parse XlsxCell object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [XlsxCell] result of parsing
    def parse(node)
      text_string_id = nil
      text_string_id = node.attribute('s').value unless node.attribute('s').nil?
      @style = CellStyle.new(parent: self).parse(text_string_id)
      if node.attribute('t').nil?
        @raw_text = node.xpath('xmlns:v').text
      else
        node.attribute('t').value == 's' ? XlsxCell.get_shared_string(node.xpath('xmlns:v').text, self) : @raw_text = node.xpath('xmlns:v').text
      end
      @formula = node.xpath('xmlns:f').text unless node.xpath('xmlns:f').text == ''
      @text = @raw_text.dup unless @raw_text.nil?
      @text.insert(0, "'") if @style.quote_prefix
      self
    end

    # Get shared string by it's number
    # @param [String] value number of shared string
    # @param [XlsxCell] cell to write value of string
    # @return [Nothing]
    def self.get_shared_string(value, cell)
      return '' if value == ''
      XLSXWorkbook.shared_strings[value.to_i].xpath('*').each do |si_node_child|
        case si_node_child.name
        when 'r'
          cell.character = ParagraphRun.new(parent: cell).parse(si_node_child)
        when 't'
          cell.raw_text = si_node_child.text
        end
      end
    end
  end
end
