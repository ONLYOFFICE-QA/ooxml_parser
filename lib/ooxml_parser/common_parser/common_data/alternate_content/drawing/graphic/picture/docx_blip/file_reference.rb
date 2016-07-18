module OoxmlParser
  # Class for storing image data
  class FileReference
    # @return [String] id of resource
    attr_accessor :resource_id
    # @return [String] path to file
    attr_accessor :path
    # @return [String] content of file
    attr_accessor :content

    def self.parse(node)
      file_ref = FileReference.new
      file_ref.resource_id = node.attribute('embed').value if node.attribute('embed')
      file_ref.resource_id = node.attribute('id').value if node.attribute('id')
      file_ref.path = OOXMLDocumentObject.get_link_from_rels(file_ref.resource_id).gsub('..', '')
      raise LoadError, "Cant find path to media file by id: #{file_ref.resource_id}" if file_ref.path.empty?
      return file_ref if file_ref.path == 'NULL'
      full_path_to_file = OOXMLDocumentObject.path_to_folder + OOXMLDocumentObject.root_subfolder + file_ref.path
      if File.exist?(full_path_to_file)
        file_ref.content = File.read(full_path_to_file)
      else
        warn "Couldn't find #{full_path_to_file} file on filesystem. Possible problem in original document"
      end
      file_ref
    end
  end
end
