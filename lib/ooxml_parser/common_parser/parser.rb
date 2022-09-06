# frozen_string_literal: true

require_relative 'parser/encryption_checker'
require_relative 'parser/ooxml_file'

module OoxmlParser
  # Basic class for OoxmlParser
  class Parser
    class << self
      # Base method to yield parse document of any type
      # @param [OoxmlFile] file with data
      # @return [CommonDocumentStructure] structure of doc
      def parse_format(file)
        return nil if EncryptionChecker.new(file.path).encrypted?

        file.copy_file_and_rename_to_zip
        file.unzip
        model = yield(file)
        model.file_path = file.path if model
        FileUtils.rm_rf(file.path_to_folder)
        model
      end

      # Base method to parse document of any type
      # @param path_to_file [String] file
      # @return [CommonDocumentStructure] structure of doc
      def parse(path_to_file, password: nil)
        file = OoxmlFile.new(path_to_file)
        file = file.decrypt(password) if password
        Parser.parse_format(file) do |yielded_file|
          format = yielded_file.format_by_folders
          case format
          when :docx
            DocumentStructure.new(unpacked_folder: yielded_file.path_to_folder).parse
          when :xlsx
            XLSXWorkbook.new(unpacked_folder: yielded_file.path_to_folder).parse
          when :pptx
            Presentation.new(unpacked_folder: yielded_file.path_to_folder).parse
          else
            warn "#{path_to_file} is a simple zip file without OOXML content"
          end
        end
      end
    end
  end
end
