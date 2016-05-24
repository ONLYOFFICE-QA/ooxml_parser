require_relative 'xlsx_data/view_model/workbook'

module OoxmlParser
  class XlsxParser
    def self.parse_xlsx(path_to_file)
      return nil if OOXMLDocumentObject.encrypted_file?(path_to_file)
      path_to_zip_file = OOXMLDocumentObject.copy_file_and_rename_to_zip(path_to_file)
      OOXMLDocumentObject.unzip_file(path_to_zip_file, path_to_zip_file.sub(File.basename(path_to_zip_file), ''))
      OOXMLDocumentObject.path_to_folder = path_to_zip_file.sub(File.basename(path_to_zip_file), '')
      xlsx_model = XLSXWorkbook.parse(OOXMLDocumentObject.path_to_folder)
      xlsx_model.file_path = path_to_file
      FileUtils.remove_dir(OOXMLDocumentObject.path_to_folder)
      xlsx_model
    end
  end
end
