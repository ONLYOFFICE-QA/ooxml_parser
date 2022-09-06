# frozen_string_literal: true

module OoxmlParser
  # Class for actions with OOXML file
  class OoxmlFile
    # @return [String] path to file
    attr_reader :path

    def initialize(path)
      @path = path
    end

    # Copy this file and rename to zip
    # @return [String] path to result zip
    def copy_file_and_rename_to_zip
      file_name = File.basename(@path)
      tmp_folder = Dir.mktmpdir('ruby-ooxml-parser')
      @zip_path = "#{tmp_folder}/#{file_name}"
      FileUtils.rm_rf(tmp_folder) if File.directory?(tmp_folder)
      FileUtils.mkdir_p(tmp_folder)
      raise "Cannot find file by path #{@path}" unless File.exist?(@path)

      FileUtils.cp path, tmp_folder
    end

    # @return [String] path to folder with zip
    def path_to_folder
      @zip_path.sub(File.basename(@zip_path), '')
    end

    # Unzip specified file
    # @return [void]
    def unzip
      Zip.warn_invalid_date = false
      Zip::File.open(@zip_path) do |zip_file|
        raise LoadError, "There is no files in zip #{@zip_path}" if zip_file.entries.empty?

        zip_file.each do |file|
          file_path = File.join(path_to_folder, file.name)
          FileUtils.mkdir_p(File.dirname(file_path))
          zip_file.extract(file, file_path) unless File.exist?(file_path)
        end
      end
    end

    # @return [Symbol] file type recognized by folder structure
    def format_by_folders
      return :docx if Dir.exist?("#{path_to_folder}/word")
      return :xlsx if Dir.exist?("#{path_to_folder}/xl")
      return :pptx if Dir.exist?("#{path_to_folder}/ppt")

      :zip
    end

    # Decrypt file protected with password
    # @param password [String] password to file
    # @return [OoxmlFile] path to decrypted file
    def decrypt(password)
      file_name = File.basename(@path)
      tmp_folder = Dir.mktmpdir('ruby-ooxml-parser')
      decrypted_path = "#{tmp_folder}/#{file_name}"
      binary_password = password.encode('utf-16le').bytes.pack('c*').encode('binary')
      OoxmlDecrypt::EncryptedFile.decrypt_to_file(@path, binary_password, decrypted_path)

      new(decrypted_path)
    end
  end
end
