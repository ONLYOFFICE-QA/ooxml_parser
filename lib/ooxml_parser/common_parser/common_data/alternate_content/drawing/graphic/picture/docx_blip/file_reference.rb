module OoxmlParser
  # Class for storing image data
  class FileReference < OOXMLDocumentObject
    # @return [String] id of resource
    attr_accessor :resource_id
    # @return [String] path to file
    attr_accessor :path
    # @return [String] content of file
    attr_accessor :content

    def parse(node)
      @resource_id = node.attribute('embed').value if node.attribute('embed')
      @resource_id = node.attribute('id').value if node.attribute('id')
      @path = OOXMLDocumentObject.get_link_from_rels(@resource_id).gsub('..', '')
      raise LoadError, "Cant find path to media file by id: #{@resource_id}" if @path.empty?
      return self if @path == 'NULL'
      full_path_to_file = OOXMLDocumentObject.path_to_folder + OOXMLDocumentObject.root_subfolder + @path
      if File.exist?(full_path_to_file)
        @content = IO.binread(full_path_to_file)
      else
        warn "Couldn't find #{full_path_to_file} file on filesystem. Possible problem in original document"
      end
      self
    end
  end
end
