# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `snd` tags
  class Sound < OOXMLDocumentObject
    # @return [String] name of sound
    attr_reader :name
    # @return [FileReference] image structure
    attr_accessor :file_reference

    def initialize(path = '', name = '', parent: nil)
      @path = path
      @name = name
      super(parent: parent)
    end

    # Parse Sound
    # @param [Nokogiri::XML:Node] node with NumberingProperties
    # @return [Sound] result of parsing
    def parse(node)
      @file_reference = FileReference.new(parent: self).parse(node)
      node.attributes.each do |key, value|
        case key
        when 'name'
          @name = value.value.to_s
        end
      end
      self
    end
  end
end
