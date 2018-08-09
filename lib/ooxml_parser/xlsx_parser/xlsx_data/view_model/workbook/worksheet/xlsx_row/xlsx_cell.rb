require_relative 'xlsx_cell/formula'
# Single Cell of XLSX
module OoxmlParser
  class XlsxCell < OOXMLDocumentObject
    attr_accessor :text, :formula, :character

    # @return [String] text without applying any style modificators, like quote_prefix
    attr_accessor :raw_text
    # @return [Integer] index of style
    attr_reader :style_index

    def initialize(parent: nil)
      @style_index = 0 # default style is zero
      @text = ''
      @raw_text = ''
      @parent = parent
    end

    # Parse XlsxCell object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [XlsxCell] result of parsing
    def parse(node)
      if node.attribute('t')
        node.attribute('t').value == 's' ? get_shared_string(node.xpath('xmlns:v').text) : @raw_text = node.xpath('xmlns:v').text
      else
        @raw_text = node.xpath('xmlns:v').text
      end
      node.attributes.each do |key, value|
        case key
        when 's'
          @style_index = value.value.to_i
        end
      end
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'f'
          @formula = Formula.new(parent: self).parse(node_child)
        end
      end
      @text = @raw_text.dup if @raw_text
      @text.insert(0, "'") if style.quote_prefix
      self
    end

    # @return [Xf] style of cell
    def style
      return nil unless @style_index
      root_object.style_sheet.cell_xfs.xf_array[@style_index]
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
