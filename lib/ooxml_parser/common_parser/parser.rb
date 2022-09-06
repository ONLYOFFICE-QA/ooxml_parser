# frozen_string_literal: true

require_relative 'parser/encryption_checker'

module OoxmlParser
  # Basic class for OoxmlParser
  class Parser
    class << self
      # Base method to yield parse document of any type
      # @param path_to_file [String] file
      # @return [CommonDocumentStructure] structure of doc
      def parse_format(path_to_file)
        return nil if EncryptionChecker.new(path_to_file).encrypted?

        path_to_zip_file = copy_file_and_rename_to_zip(path_to_file)
        path_to_folder = path_to_zip_file.sub(File.basename(path_to_zip_file), '')
        unzip_file(path_to_zip_file, path_to_folder)
        model = yield(path_to_folder)
        model.file_path = path_to_file if model
        FileUtils.rm_rf(path_to_folder)
        model
      end

      # Base method to parse document of any type
      # @param path_to_file [String] file
      # @return [CommonDocumentStructure] structure of doc
      def parse(path_to_file, password: nil)
        path_to_file = decrypt_file(path_to_file, password) if password
        Parser.parse_format(path_to_file) do |path_to_folder|
          format = Parser.recognize_folder_format(path_to_folder)
          case format
          when :docx
            DocumentStructure.new(unpacked_folder: path_to_folder).parse
          when :xlsx
            XLSXWorkbook.new(unpacked_folder: path_to_folder).parse
          when :pptx
            Presentation.new(unpacked_folder: path_to_folder).parse
          else
            warn "#{path_to_file} is a simple zip file without OOXML content"
          end
        end
      end

      # Recognize folder format
      # @param directory [String] path to dirctory
      # @return [Symbol] type of document
      def recognize_folder_format(directory)
        return :docx if Dir.exist?("#{directory}/word")
        return :xlsx if Dir.exist?("#{directory}/xl")
        return :pptx if Dir.exist?("#{directory}/ppt")
      end

      # Copy this file and rename to zip
      # @param path [String] path to file
      # @return [String] path to result zip
      def copy_file_and_rename_to_zip(path)
        file_name = File.basename(path)
        tmp_folder = Dir.mktmpdir('ruby-ooxml-parser')
        file_path = "#{tmp_folder}/#{file_name}"
        FileUtils.rm_rf(tmp_folder) if File.directory?(tmp_folder)
        FileUtils.mkdir_p(tmp_folder)
        raise "Cannot find file by path #{path}" unless File.exist?(path)

        FileUtils.cp path, tmp_folder
        file_path
      end

      # Decrypt file protected with password
      # @param path [String] path to file
      # @param password [String] password to file
      # @return [String] path to decrypted file
      def decrypt_file(path, password)
        file_name = File.basename(path)
        tmp_folder = Dir.mktmpdir('ruby-ooxml-parser')
        decrypted_path = "#{tmp_folder}/#{file_name}"
        binary_password = password.encode('utf-16le').bytes.pack('c*').encode('binary')
        OoxmlDecrypt::EncryptedFile.decrypt_to_file(path, binary_password, decrypted_path)

        decrypted_path
      end

      # Unzip specified file
      # @param path_to_file [String] path to zip file
      # @param destination [String] folder to extract
      # @return [void]
      def unzip_file(path_to_file, destination)
        Zip.warn_invalid_date = false
        Zip::File.open(path_to_file) do |zip_file|
          raise LoadError, "There is no files in zip #{path_to_file}" if zip_file.entries.empty?

          zip_file.each do |file|
            file_path = File.join(destination, file.name)
            FileUtils.mkdir_p(File.dirname(file_path))
            zip_file.extract(file, file_path) unless File.exist?(file_path)
          end
        end
      end
    end
  end
end
