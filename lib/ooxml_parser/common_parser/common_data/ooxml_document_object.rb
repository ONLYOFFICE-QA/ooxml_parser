require 'filemagic'
require 'securerandom'
require 'nokogiri'
require 'zip'
require_relative 'ooxml_document_object/ooxml_document_object_helper'
require_relative 'ooxml_document_object/ooxml_object_attribute_helper'

module OoxmlParser
  class OOXMLDocumentObject
    include OoxmlDocumentObjectHelper
    include OoxmlObjectAttributeHelper
    # @return [OOXMLDocumentObject] object which hold current object
    attr_accessor :parent

    def initialize(parent: nil)
      @parent = parent
    end

    def ==(other)
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
      Nokogiri::XML(File.open(xml_path), &:strict)
    end

    class << self
      attr_accessor :root_subfolder
      attr_accessor :theme
      attr_accessor :xmls_stack
      attr_accessor :path_to_folder

      # @param path_to_file [String] file
      # @return [True, False] Check if file is protected by password on open
      def encrypted_file?(path_to_file)
        if Gem.win_platform?
          warn 'FileMagic and checking file for encryption is not supported on Windows'
          return false
        end
        file_result = FileMagic.new(:mime).file(path_to_file)
        # Support of Encrtypted status in `file` util was introduced in file v5.20
        # but LTS version of ubuntu before 16.04 uses older `file` and it return `Composite Document`
        # https://github.com/file/file/blob/master/ChangeLog#L217
        if file_result.include?('encrypted') ||
           file_result.include?('Composite Document File V2 Document, No summary info') ||
           file_result.include?('application/CDFV2-corrupt')
          warn("File #{path_to_file} is encrypted. Can't parse it")
          return true
        end
        false
      end

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

      def dir
        OOXMLDocumentObject.path_to_folder + File.dirname(OOXMLDocumentObject.xmls_stack.last) + '/'
      end

      def current_xml
        OOXMLDocumentObject.path_to_folder + OOXMLDocumentObject.xmls_stack.last
      end

      def add_to_xmls_stack(path)
        OOXMLDocumentObject.xmls_stack << if path.include?('..')
                                            "#{File.dirname(OOXMLDocumentObject.xmls_stack.last)}/#{path}"
                                          elsif path.start_with?(OOXMLDocumentObject.root_subfolder)
                                            path
                                          else
                                            OOXMLDocumentObject.root_subfolder + path
                                          end
      end

      def get_link_from_rels(id)
        rels_path = dir + "_rels/#{File.basename(OOXMLDocumentObject.xmls_stack.last)}.rels"
        raise LoadError, "Cannot find .rels file by path: #{rels_path}" unless File.exist?(rels_path)

        relationships = Relationships.new.parse_file(rels_path)
        relationships.target_by_id(id)
      end
    end
  end
end
