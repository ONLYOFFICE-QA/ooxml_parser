# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `snd` tags
  class Sound < OOXMLDocumentObject
    attr_accessor :name
    # @return [FileReference] image structure
    attr_accessor :file_reference

    def initialize(path = '', name = '', parent: nil)
      @path = path
      @name = name
      @parent = parent
    end

    # Parse Sound
    # @param [Nokogiri::XML:Node] node with NumberingProperties
    # @return [Sound] result of parsing
    def parse(node)
      @file_reference = FileReference.new(parent: self).parse(node)
      @name = node.attribute('name').value
      self
    end
  end
end
