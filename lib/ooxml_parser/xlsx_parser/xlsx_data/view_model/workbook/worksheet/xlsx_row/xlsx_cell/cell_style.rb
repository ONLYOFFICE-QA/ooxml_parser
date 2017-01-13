require_relative 'cell_style/foreground_color'
require_relative 'cell_style/alignment'
require_relative 'cell_style/ooxml_font'
module OoxmlParser
  # Class for parsing cell style
  class CellStyle < OOXMLDocumentObject
    ALL_FORMAT_VALUE = %w|0 0.00 #,##0 #,##0.00 $#,##0_);($#,##0) $#,##0_);[Red]($#,##0) $#,##0.00_);($#,##0.00)
                          $#,##0.00_);[Red]($#,##0.00)
                          0.00% 0.00%
                          0.00E+00
                          #\ ?/?
                          #\ ??/??
                          m/d/yyyy
                          d-mmm-yy
                          d-mmm
                          mmm-yy
                          h:mm AM/PM
                          h:mm:ss AM/PM
                          h:mm
                          h:mm:ss
                          m/d/yyyy h:mm
                          0
                          0
                          0
                          0
                          0
                          0
                          0
                          0
                          0
                          0
                          0
                          0
                          0
                          0
                          #,##0_);(#,##0)
                          #,##0_);[Red](#,##0)
                          #,##0.00_);(#,##0.00)
                          #,##0.00_);[Red](#,##0.00)
                          0
                          0
                          0
                          0
                          mm:ss
                          General
                          mm:ss.0
                          ##0.0E+0
                          @|.freeze

    attr_accessor :font, :borders, :fill_color, :numerical_format, :alignment
    # @return [True, False] check if style should add QuotePrefix (' symbol) to start of the string
    attr_accessor :quote_prefix
    # @return [True, False] is number format is applied
    attr_accessor :apply_number_format

    def initialize(font = nil,
                   borders = nil,
                   fill_color = ForegroundColor.new,
                   numerical_format = 'General',
                   alignment = XlsxAlignment.new)
      @font = font
      @borders = borders
      @fill_color = fill_color
      @numerical_format = numerical_format
      @alignment = alignment
    end

    # Parse CellStyle object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [CellStyle] result of parsing
    def parse(node)
      current_cell_style = XLSXWorkbook.styles_node.xpath('//xmlns:cellXfs/xmlns:xf')[node.to_i]
      @font = if current_cell_style.attribute('applyFont').nil? || current_cell_style.attribute('applyFont').value == '0'
                OOXMLFont.new(parent: self).parse(0)
              else
                OOXMLFont.new(parent: self).parse(current_cell_style.attribute('fontId').value.to_i)
              end
      unless current_cell_style.attribute('applyBorder').nil? || current_cell_style.attribute('applyBorder').value == '0'
        @borders = Borders.parse_from_style(current_cell_style.attribute('borderId').value.to_i)
      end
      unless current_cell_style.attribute('applyFill').nil? || current_cell_style.attribute('applyFill').value == '0'
        @fill_color = ForegroundColor.parse(current_cell_style.attribute('fillId').value.to_i)
      end
      unless current_cell_style.attribute('applyNumberFormat').nil? || current_cell_style.attribute('applyNumberFormat').value == '0'
        @apply_number_format = true
        format_id = current_cell_style.attribute('numFmtId').value.to_i
        XLSXWorkbook.styles_node.xpath('//xmlns:numFmt').each do |numeric_format|
          if format_id == numeric_format.attribute('numFmtId').value.to_i
            @numerical_format = numeric_format.attribute('formatCode').value
          elsif CellStyle::ALL_FORMAT_VALUE[format_id - 1]
            @numerical_format = CellStyle::ALL_FORMAT_VALUE[format_id - 1]
          end
        end
        @numerical_format = CellStyle::ALL_FORMAT_VALUE[format_id - 1] if CellStyle::ALL_FORMAT_VALUE[format_id - 1]
      end
      unless current_cell_style.attribute('applyAlignment').nil? || current_cell_style.attribute('applyAlignment').value == '0'
        current_cell_style.xpath('*').each do |node_child|
          case node_child.name
          when 'alignment'
            @alignment.parse(node_child)
          end
        end
      end
      @quote_prefix = option_enabled?(current_cell_style, 'quotePrefix')
      self
    end
  end
end
