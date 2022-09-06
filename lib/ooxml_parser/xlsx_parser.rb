# frozen_string_literal: true

require_relative 'xlsx_parser/workbook'

module OoxmlParser
  # Basic class for parsing xlsx
  class XlsxParser
    # Parse xlsx file
    # @param path_to_file [String] file path
    # @return [XLSXWorkbook] result of parse
    def self.parse_xlsx(path_to_file)
      file = OoxmlFile.new(path_to_file)
      Parser.parse_format(file) do |yielded_file|
        XLSXWorkbook.new(unpacked_folder: yielded_file.path_to_folder).parse
      end
    end
  end
end
