require_relative 'xlsx_data/view_model/workbook'

module OoxmlParser
  class XlsxParser
    def self.parse_xlsx(path_to_file)
      Parser.parse(path_to_file) do
        XLSXWorkbook.parse
      end
    end
  end
end
