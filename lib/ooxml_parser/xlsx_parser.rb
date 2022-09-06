# frozen_string_literal: true

require_relative 'xlsx_parser/workbook'

module OoxmlParser
  # Basic class for parsing xlsx
  class XlsxParser
    # Parse xlsx file
    # @param path_to_file [String] file path
    # @return [XLSXWorkbook] result of parse
    def self.parse_xlsx(path_to_file)
      Parser.parse_format(path_to_file) do |file|
        XLSXWorkbook.new(unpacked_folder: file.path_to_folder).parse
      end
    end
  end
end
