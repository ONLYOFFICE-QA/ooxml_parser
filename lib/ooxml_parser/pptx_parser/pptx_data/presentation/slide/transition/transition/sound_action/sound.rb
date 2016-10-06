module OoxmlParser
  class Sound < OOXMLDocumentObject
    attr_accessor :name
    # @return [FileReference] image structure
    attr_accessor :file_reference

    def initialize(path = '', name = '')
      @path = path
      @name = name
    end

    def self.parse(sound_node)
      sound = Sound.new
      sound.file_reference = FileReference.new(parent: self).parse(sound_node)
      sound.name = sound_node.attribute('name').value
      sound
    end
  end
end
