# frozen_string_literal: true

require_relative 'xlsx_cell/formula'
module OoxmlParser
  # Single Cell of XLSX
  class XlsxCell < OOXMLDocumentObject
    attr_accessor :formula, :character

    # @return [String] text without applying any style modificators, like quote_prefix
    attr_accessor :raw_text
    # @return [Integer] index of style
    attr_reader :style_index
    # @return [String] value of cell
    attr_reader :value
    # @return [String] type of string
    attr_reader :type

    def initialize(parent: nil)
      @style_index = 0 # default style is zero
      @raw_text = ''
      @parent = parent
    end

    # Parse XlsxCell object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [XlsxCell] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 's'
          @style_index = value.value.to_i
        when 't'
          @type = value.value.to_s
        end
      end
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'f'
          @formula = Formula.new(parent: self).parse(node_child)
        when 'v'
          @value = TextValue.new(parent: self).parse(node_child)
        end
      end
      parse_text_data
      self
    end

    # @return [Xf] style of cell
    def style
      return nil unless @style_index

      root_object.style_sheet.cell_xfs.xf_array[@style_index]
    end

    # @return [String] text with modifiers
    def text
      return '' unless @raw_text
      return @raw_text.dup.insert(0, "'") if style.quote_prefix

      @raw_text
    end

    private

    # @return [Nothing] parse text data
    def parse_text_data
      if @type && @value
        type == 's' ? get_shared_string(value.value) : @raw_text = value.value
      elsif @value
        @raw_text = value.value if @value
      end
    end

    # Get shared string by it's number
    # @return [Nothing]
    def get_shared_string(value)
      return '' if value == ''

      string = root_object.shared_strings_table.string_indexes[value.to_i]
      @character = string.run
      @raw_text = string.text
    end
  end
end
