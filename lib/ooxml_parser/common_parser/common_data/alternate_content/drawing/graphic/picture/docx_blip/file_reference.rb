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

      full_path_to_file = OOXMLDocumentObject.path_to_folder + OOXMLDocumentObject.root_subfolder + @path.gsub('..', '')
      if File.exist?(full_path_to_file)
        @content = IO.binread(full_path_to_file)
      else
        warn "Couldn't find #{full_path_to_file} file on filesystem. Possible problem in original document"
      end
      self
    end
  end
end
