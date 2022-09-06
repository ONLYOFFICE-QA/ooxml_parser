# frozen_string_literal: true

require 'securerandom'
require 'nokogiri'
require 'zip'
require 'ooxml_decrypt'
require_relative 'ooxml_document_object/nokogiri_parsing_exception'
require_relative 'ooxml_document_object/ooxml_document_object_helper'
require_relative 'ooxml_document_object/ooxml_object_attribute_helper'

module OoxmlParser
  # Basic class for any OOXML Document Object
  class OOXMLDocumentObject
    include OoxmlDocumentObjectHelper
    include OoxmlObjectAttributeHelper
    # @return [OOXMLDocumentObject] object which hold current object
    attr_accessor :parent

    def initialize(parent: nil)
      @parent = parent
    end

    # Compare this object to other
    # @param other [Object] any other object
    # @return [True, False] result of comparision
    def ==(other)
      return false if self.class != other.class

      instance_variables.each do |current_attribute|
        next if current_attribute == :@parent
        return false unless instance_variable_get(current_attribute) == other.instance_variable_get(current_attribute)
      end
      true
    end

    # @return [True, false] if structure contain any user data
    def with_data?
      true
    end

    # @return [Nokogiri::XML::Document] result of parsing xml via nokogiri
    def parse_xml(xml_path)
      xml = Nokogiri::XML(File.open(xml_path), &:strict)
      unless xml.errors.empty?
        raise NokogiriParsingException,
              "Nokogiri found errors in file: #{xml_path}. Errors: #{xml.errors}"
      end
      xml
    end

    class << self
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
