require_relative 'xlsx_cell/cell_style'
# Single Cell of XLSX
module OoxmlParser
  class XlsxCell < OOXMLDocumentObject
    attr_accessor :style, :text, :formula, :character

    # @return [String] text without applying any style modificators, like quote_prefix
    attr_accessor :raw_text

    def initialize(style = nil, text = '')
      @style = style
      @text = text
      @raw_text = ''
    end

    def self.parse(cell_node)
      text_string_id = nil
      text_string_id = cell_node.attribute('s').value unless cell_node.attribute('s').nil?
      cell = XlsxCell.new(CellStyle.parse(text_string_id))
      if cell_node.attribute('t').nil?
        cell.raw_text = cell_node.xpath('xmlns:v').text
      else
        cell_node.attribute('t').value == 's' ? get_shared_string(cell_node.xpath('xmlns:v').text, cell) : cell.raw_text = cell_node.xpath('xmlns:v').text
      end
      cell.formula = cell_node.xpath('xmlns:f').text unless cell_node.xpath('xmlns:f').text == ''
      cell.text = cell.raw_text.dup unless cell.raw_text.nil?
      cell.text.insert(0, "'") if cell.style.quote_prefix
      cell
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
          cell.character = ParagraphRun.parse(si_node_child)
        when 't'
          cell.raw_text = si_node_child.text
        end
      end
    end
  end
end
