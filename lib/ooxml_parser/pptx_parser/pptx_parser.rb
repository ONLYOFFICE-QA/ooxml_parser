require_relative 'pptx_data/presentation.rb'

module OoxmlParser
  class PptxParser
    def self.parse_pptx(path_to_file)
      path_to_zip_file = OOXMLDocumentObject.copy_file_and_rename_to_zip(path_to_file)
      OOXMLDocumentObject.unzip_file(path_to_zip_file, path_to_zip_file.sub(File.basename(path_to_zip_file), ''))
      OOXMLDocumentObject.path_to_folder = path_to_zip_file.sub(File.basename(path_to_zip_file), '')
      pptx_model = Presentation.parse(OOXMLDocumentObject.path_to_folder)
      pptx_model.file_path = path_to_file
      FileUtils.remove_dir(OOXMLDocumentObject.path_to_folder)
      pptx_model
    end
  end
end
