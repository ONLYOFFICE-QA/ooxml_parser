# frozen_string_literal: true

require_relative 'parser/encryption_checker'

module OoxmlParser
  # Basic class for OoxmlParser
  class Parser
    # Base method to yield parse document of any type
    # @param path_to_file [String] file
    # @return [CommonDocumentStructure] structure of doc
    def self.parse_format(path_to_file)
      return nil if EncryptionChecker.new(path_to_file).encrypted?

      path_to_zip_file = OOXMLDocumentObject.copy_file_and_rename_to_zip(path_to_file)
      OOXMLDocumentObject.path_to_folder = path_to_zip_file.sub(File.basename(path_to_zip_file), '')
      OOXMLDocumentObject.unzip_file(path_to_zip_file, OOXMLDocumentObject.path_to_folder)
      model = yield
      model.file_path = path_to_file if model
      FileUtils.rm_rf(OOXMLDocumentObject.path_to_folder)
      model
    end

    # Base method to parse document of any type
    # @param path_to_file [String] file
    # @return [CommonDocumentStructure] structure of doc
    def self.parse(path_to_file, password: nil)
      path_to_file = OOXMLDocumentObject.decrypt_file(path_to_file, password) if password
      Parser.parse_format(path_to_file) do
        format = Parser.recognize_folder_format
        case format
        when :docx
          DocumentStructure.parse
        when :xlsx
          XLSXWorkbook.new.parse
        when :pptx
          Presentation.new.parse
        else
          warn "#{path_to_file} is a simple zip file without OOXML content"
        end
      end
    end

    # Recognize folder format
    # @param directory [String] path to dirctory
    # @return [Symbol] type of document
    def self.recognize_folder_format(directory = OOXMLDocumentObject.path_to_folder)
      return :docx if Dir.exist?("#{directory}/word")
      return :xlsx if Dir.exist?("#{directory}/xl")
      return :pptx if Dir.exist?("#{directory}/ppt")
    end
  end
end
