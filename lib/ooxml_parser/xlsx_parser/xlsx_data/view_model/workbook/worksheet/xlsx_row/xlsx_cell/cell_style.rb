require_relative 'cell_style/foreground_color'
require_relative 'cell_style/alignment'
require_relative 'cell_style/ooxml_font'
module OoxmlParser
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
                          @|

    attr_accessor :font, :borders, :fill_color, :numerical_format, :alignment
    # @return [True, False] check if style should add QuotePrefix (' symbol) to start of the string
    attr_accessor :quote_prefix

    def initialize(font = nil, borders = nil, fill_color = ForegroundColor.new, numerical_format = 'General', alignment = nil)
      @font = font
      @borders = borders
      @fill_color = fill_color
      @numerical_format = numerical_format
      @alignment = alignment
    end

    def self.parse(style_number)
      current_cell_style = XLSXWorkbook.styles_node.xpath('//xmlns:cellXfs/xmlns:xf')[style_number.to_i]
      cell_style = CellStyle.new
      cell_style.alignment = XlsxAlignment.new
      if current_cell_style.attribute('applyFont').nil? || current_cell_style.attribute('applyFont').value == '0'
        cell_style.font = OOXMLFont.parse(0)
      else
        cell_style.font = OOXMLFont.parse(current_cell_style.attribute('fontId').value.to_i)
      end
      unless current_cell_style.attribute('applyBorder').nil? || current_cell_style.attribute('applyBorder').value == '0'
        cell_style.borders = Borders.parse_from_style(current_cell_style.attribute('borderId').value.to_i)
      end
      unless current_cell_style.attribute('applyFill').nil? || current_cell_style.attribute('applyFill').value == '0'
        cell_style.fill_color = ForegroundColor.parse(current_cell_style.attribute('fillId').value.to_i)
      end
      unless current_cell_style.attribute('applyNumberFormat').nil? || current_cell_style.attribute('applyNumberFormat').value == '0'
        format_id = current_cell_style.attribute('numFmtId').value.to_i
        XLSXWorkbook.styles_node.xpath('//xmlns:numFmt').each do |numeric_format|
          if format_id == numeric_format.attribute('numFmtId').value.to_i
            cell_style.numerical_format = numeric_format.attribute('formatCode').value
          elsif CellStyle::ALL_FORMAT_VALUE[format_id - 1]
            cell_style.numerical_format = CellStyle::ALL_FORMAT_VALUE[format_id - 1]
          end
        end
        cell_style.numerical_format = CellStyle::ALL_FORMAT_VALUE[format_id - 1] if CellStyle::ALL_FORMAT_VALUE[format_id - 1]
      end
      unless current_cell_style.attribute('applyAlignment').nil? || current_cell_style.attribute('applyAlignment').value == '0'
        alignment_node = current_cell_style.xpath('xmlns:alignment').first
        cell_style.alignment = XlsxAlignment.parse(alignment_node) unless alignment_node.nil?
      end
      cell_style.quote_prefix = option_enabled?(current_cell_style, 'quotePrefix')
      cell_style
    end
  end
end
