module OoxmlParser
  class Parser
    # Base method to parse document of any type
    # @param path_to_file [String] file
    # @return [CommonDocumentStructure] structure of doc
    def self.parse(path_to_file)
      return nil if OOXMLDocumentObject.encrypted_file?(path_to_file)
      path_to_zip_file = OOXMLDocumentObject.copy_file_and_rename_to_zip(path_to_file)
      OOXMLDocumentObject.path_to_folder = path_to_zip_file.sub(File.basename(path_to_zip_file), '')
      OOXMLDocumentObject.unzip_file(path_to_zip_file, OOXMLDocumentObject.path_to_folder)
      model = yield
      model.file_path = path_to_file
      FileUtils.remove_dir(OOXMLDocumentObject.path_to_folder)
      model
    end
  end
end
