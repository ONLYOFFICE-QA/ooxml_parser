require_relative 'paragraph_run/run_properties'
module OoxmlParser
  class ParagraphRun < OOXMLDocumentObject
    attr_accessor :properties, :text

    def initialize(properties = RunProperties.new, text = '')
      @properties = properties
      @text = text
    end

    def self.parse(character_node)
      character = ParagraphRun.new
      character_node.xpath('*').each do |character_node_child|
        case character_node_child.name
        when 'rPr'
          character.properties = RunProperties.new(parent: character).parse(character_node_child)
        when 't'
          character.text = character_node_child.text
        end
      end
      character
    end
  end
end
