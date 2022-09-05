# frozen_string_literal: true

require 'uri'

module OoxmlParser
  # Class for storing image data
  class FileReference < OOXMLDocumentObject
    # @return [String] id of resource
    attr_accessor :resource_id
    # @return [String] path to file
    attr_accessor :path
    # @return [String] content of file
    attr_accessor :content

    # Parse FileReference object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [FileReference] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'embed', 'id', 'link'
          @resource_id = value.value
        end
      end
      return self unless @resource_id
      return self if @resource_id.empty?

      @path = OOXMLDocumentObject.get_link_from_rels(@resource_id)
      if !@path || @path.empty?
        warn "Cant find path to media file by id: #{@resource_id}"
        return self
      end
      return self if @path == 'NULL'
      return self if @path.match?(URI::DEFAULT_PARSER.make_regexp)

      full_path_to_file = OOXMLDocumentObject.path_to_folder + OOXMLDocumentObject.root_subfolder + @path.gsub('..', '')
      if File.exist?(full_path_to_file)
        @content = if File.extname(@path) == '.xlsx'
                     parse_ole_xlsx(full_path_to_file)
                   else
                     File.binread(full_path_to_file)
                   end
      else
        warn "Couldn't find #{full_path_to_file} file on filesystem. Possible problem in original document"
      end
      self
    end

    private

    # Parse ole xlsx file
    # @param [String] full_path to file
    # @return [XLSXWorkbook]
    def parse_ole_xlsx(full_path)
      # TODO: Fix this ugly hack with global vars
      #   by replacing all global variables
      stack = OOXMLDocumentObject.xmls_stack
      dir = OOXMLDocumentObject.path_to_folder
      result = OoxmlParser::Parser.parse(full_path)
      OOXMLDocumentObject.xmls_stack = stack
      OOXMLDocumentObject.path_to_folder = dir
      result
    end
  end
end
