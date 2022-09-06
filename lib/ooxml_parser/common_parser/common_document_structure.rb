# frozen_string_literal: true

require_relative 'common_data/content_types'
module OoxmlParser
  # Common document structure for DOCX, XLSX, PPTX file
  class CommonDocumentStructure < OOXMLDocumentObject
    # @return [String] path to original file
    attr_accessor :file_path
    # @return [Integer] default font size
    attr_reader :default_font_size
    # @return [Integer] default font typeface
    attr_reader :default_font_typeface
    # @return [FontStyle] Default font style of presentation
    attr_accessor :default_font_style
    # @return [ContentTypes] data about content types
    attr_accessor :content_types
    # @return [String] root sub-folder for object
    attr_reader :root_subfolder
    # @return [String] path to folder with unpacked document
    attr_reader :unpacked_folder
    # @return [Array<String>] list of xmls to parse
    attr_accessor :xmls_stack

    def initialize(params = {})
      @default_font_size = params.fetch(:default_font_size, 18)
      @default_font_typeface = params.fetch(:default_font_typeface, 'Arial')
      @default_font_style = FontStyle.new
      @unpacked_folder = params.fetch(:unpacked_folder, nil)
      @xmls_stack = []
      super(parent: nil)
    end

    # @return [String] path to current xml file
    def current_xml
      root_object.unpacked_folder + @xmls_stack.last
    end

    # Add file to parsing stack
    # @param path [String] path of file to add to stack
    # @return [void]
    def add_to_xmls_stack(path)
      @xmls_stack << if path.include?('..')
                       "#{File.dirname(@xmls_stack.last)}/#{path}"
                     elsif path.start_with?(@root_subfolder)
                       path
                     else
                       @root_subfolder + path
                     end
    end

    # Get link to file from rels file
    # @param id [String] file to get
    # @return [String] result
    def get_link_from_rels(id)
      dir = "#{unpacked_folder}#{File.dirname(@xmls_stack.last)}/"
      rels_path = dir + "_rels/#{File.basename(@xmls_stack.last)}.rels"
      raise LoadError, "Cannot find .rels file by path: #{rels_path}" unless File.exist?(rels_path)

      relationships = Relationships.new.parse_file(rels_path)
      relationships.target_by_id(id)
    end
  end
end
