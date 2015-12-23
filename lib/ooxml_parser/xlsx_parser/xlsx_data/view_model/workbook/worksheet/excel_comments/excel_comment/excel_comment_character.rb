require_relative 'excel_comment_character/excel_comment_character_properties'
# Single Properties Character (Run) of XLSX comment
module OoxmlParser
  class ExcelCommentCharacter
    attr_accessor :properties, :text

    def initialize(properties = nil, text = '')
      @properties = properties
      @text = text
    end

    def self.parse(character_node)
      character = ExcelCommentCharacter.new
      character_node.xpath('*').each do |character_node_child|
        case character_node_child.name
        when 'rPr'
          character.properties = ExcelCommentCharacterProperties.parse(character_node_child)
        when 't'
          character.text = character_node_child.text
        end
      end
      character
    end
  end
end
