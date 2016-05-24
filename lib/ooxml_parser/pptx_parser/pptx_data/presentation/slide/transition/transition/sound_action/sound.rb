module OoxmlParser
  class Sound < OOXMLDocumentObject
    attr_accessor :path, :name

    def initialize(path = '', name = '')
      @path = path
      @name = name
    end

    def self.parse(sound_node)
      sound = Sound.new
      sound.path = OOXMLDocumentObject.copy_media_file("#{OOXMLDocumentObject.root_subfolder}/#{get_link_from_rels(sound_node.attribute('embed').value).gsub('..', '')}")
      sound.name = sound_node.attribute('name').value
      sound
    end
  end
end
