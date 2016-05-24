require 'securerandom'
require 'nokogiri'
require 'xmlsimple'
require 'zip'

module OoxmlParser
  class OOXMLDocumentObject
    DEFAULT_DIRECTORY_FOR_MEDIA = '/tmp'.freeze

    def ==(other)
      instance_variables.each do |current_attribute|
        unless instance_variable_get(current_attribute) == other.instance_variable_get(current_attribute)
          return false
        end
      end
      true
    end

    class << self
      attr_accessor :namespace_prefix
      attr_accessor :root_subfolder
      attr_accessor :theme
      attr_accessor :xmls_stack
      attr_accessor :path_to_folder

      # @param path_to_file [String] file
      # @return [True, False] Check if file is protected by password on open
      def encrypted_file?(path_to_file)
        file_result = `file #{path_to_file}`
        # Support of Encrtypted status in `file` util was introduced in file v5.20
        # but LTS version of ubuntu before 16.04 uses older `file` and it return `Composite Document`
        # https://github.com/file/file/blob/master/ChangeLog#L217
        if file_result.include?('Encrypted') || file_result.include?('Composite Document File V2 Document, No summary info')
          warn("File #{path_to_file} is encrypted. Can't parse it")
          return true
        end
        false
      end

      def copy_file_and_rename_to_zip(path)
        file_name = File.basename(path)
        tmp_folder = "/tmp/office_open_xml_parser_#{SecureRandom.uuid}"
        file_path = "#{tmp_folder}/#{file_name}"
        FileUtils.rm_rf(tmp_folder) if File.directory?(tmp_folder)
        FileUtils.mkdir_p(tmp_folder)
        path = "#{Dir.pwd}/#{path}" unless path[0] == '/'
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
                                          else
                                            path
                                          end
      end

      def get_link_from_rels(id)
        rels_path = dir + "_rels/#{File.basename(OOXMLDocumentObject.xmls_stack.last)}.rels"
        raise LoadError, "Cannot find .rels file by path: #{rels_path}" unless File.exist?(rels_path)
        relationships = XmlSimple.xml_in(File.open(rels_path))
        relationships['Relationship'].each { |relationship| return relationship['Target'] if id == relationship['Id'] }
      end

      def media_folder
        path = "#{DEFAULT_DIRECTORY_FOR_MEDIA}/media_from_#{@file_name}"
        FileUtils.mkdir(path) unless File.exist?(path)
        path + '/'
      end

      def option_enabled?(node, attribute_name = 'val')
        return true if node.to_s == '1'
        return false if node.to_s == '0'
        return false if node.attribute(attribute_name).nil?
        status = node.attribute(attribute_name).value
        status == 'true' || status == 'on' || status == '1'
      end

      def copy_media_file(path_to_file)
        folder_to_save_media = '/tmp/media_from_' + File.basename(OOXMLDocumentObject.path_to_folder)
        path_to_copied_file = folder_to_save_media + '/' + File.basename(path_to_file)
        FileUtils.mkdir(folder_to_save_media) unless File.exist?(folder_to_save_media)
        FileUtils.copy_file(OOXMLDocumentObject.path_to_folder + path_to_file, path_to_copied_file)
        path_to_copied_file
      end
    end
  end
end
