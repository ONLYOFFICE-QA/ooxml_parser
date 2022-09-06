# frozen_string_literal: true

require_relative 'parser/encryption_checker'
require_relative 'parser/ooxml_file'

module OoxmlParser
  # Basic class for OoxmlParser
  class Parser
    class << self
      # Base method to yield parse document of any type
      # @param path_to_file [String] file
      # @return [CommonDocumentStructure] structure of doc
      def parse_format(path_to_file)
        return nil if EncryptionChecker.new(path_to_file).encrypted?

        file = OoxmlFile.new(path_to_file)
        file.copy_file_and_rename_to_zip
        file.unzip
        model = yield(file)
        model.file_path = path_to_file if model
        FileUtils.rm_rf(file.path_to_folder)
        model
      end

      # Base method to parse document of any type
      # @param path_to_file [String] file
      # @return [CommonDocumentStructure] structure of doc
      def parse(path_to_file, password: nil)
        Parser.parse_format(path_to_file) do |file|
          file = file.decrypt(password) if password
          format = file.format_by_folders
          case format
          when :docx
            DocumentStructure.new(unpacked_folder: file.path_to_folder).parse
          when :xlsx
            XLSXWorkbook.new(unpacked_folder: file.path_to_folder).parse
          when :pptx
            Presentation.new(unpacked_folder: file.path_to_folder).parse
          else
            warn "#{path_to_file} is a simple zip file without OOXML content"
          end
        end
      end
    end
  end
end
