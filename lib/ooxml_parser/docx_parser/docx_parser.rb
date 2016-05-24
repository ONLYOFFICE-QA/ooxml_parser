require_relative '../common_parser/common_data/ooxml_document_object'
require_relative 'docx_data/document_structure'

module OoxmlParser
  class DocxParser
    def self.parse_docx(path_to_file)
      return nil if OOXMLDocumentObject.encrypted_file?(path_to_file)
      path_to_zip_file = OOXMLDocumentObject.copy_file_and_rename_to_zip(path_to_file)
      OOXMLDocumentObject.unzip_file(path_to_zip_file, path_to_zip_file.sub(File.basename(path_to_zip_file), ''))
      OOXMLDocumentObject.path_to_folder = path_to_zip_file.sub(File.basename(path_to_zip_file), '')
      docx_model = DocumentStructure.parse(OOXMLDocumentObject.path_to_folder)
      docx_model.file_path = path_to_file
      FileUtils.remove_dir(OOXMLDocumentObject.path_to_folder)
      docx_model
    end
  end
end
