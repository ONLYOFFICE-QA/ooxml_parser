module OoxmlParser
  class Sound < OOXMLDocumentObject
    attr_accessor :path, :name

    def initialize(path = '', name = '')
      @path = path
      @name = name
    end

    def self.parse(sound_node)
      sound = Sound.new
      path_to_original_file = dir + OOXMLDocumentObject.get_link_from_rels(sound_node.attribute('embed').value)
      sound.path = FileHelper.copy_file(path_to_original_file, "#{DEFAULT_DIRECTORY_FOR_MEDIA}/media_from_#{File.basename(OOXMLDocumentObject.path_to_folder)}/")
      sound.name = sound_node.attribute('name').value
      sound
    end
  end
end
