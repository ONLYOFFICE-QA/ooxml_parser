# frozen_string_literal: true

require_relative 'xlsx_data/view_model/workbook'

module OoxmlParser
  # Basic class for parsing xlsx
  class XlsxParser
    def self.parse_xlsx(path_to_file)
      Parser.parse_format(path_to_file) do
        XLSXWorkbook.new.parse
      end
    end
  end
end
