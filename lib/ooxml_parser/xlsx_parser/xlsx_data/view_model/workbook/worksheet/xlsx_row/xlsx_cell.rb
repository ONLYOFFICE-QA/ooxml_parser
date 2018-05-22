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
      text_string_id = node.attribute('s').value if node.attribute('s')
      @style = CellStyle.new(parent: self).parse(text_string_id)
      if node.attribute('t')
        node.attribute('t').value == 's' ? get_shared_string(node.xpath('xmlns:v').text) : @raw_text = node.xpath('xmlns:v').text
      else
        @raw_text = node.xpath('xmlns:v').text
      end
      @formula = node.xpath('xmlns:f').text unless node.xpath('xmlns:f').text == ''
      @text = @raw_text.dup unless @raw_text.nil?
      @text.insert(0, "'") if @style.quote_prefix
      self
    end

    private

    # Get shared string by it's number
    # @param [String] value number of shared string
    # @return [Nothing]
    def get_shared_string(value)
      return '' if value == ''
      string = root_object.shared_strings_table.string_indexes[value.to_i]
      @character = string.run
      @raw_text = string.text
    end
  end
end
