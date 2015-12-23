require_relative 'excel_comment/excel_comment_character'
# Single Comment of XLSX
module OoxmlParser
  class ExcelComment < OOXMLDocumentObject
    attr_accessor :characters

    def initialize(characters = [])
      @characters = characters
    end

    def self.parse(comment_node)
      comment = ExcelComment.new
      comment_node.xpath('xmlns:text/xmlns:r').each do |character_node|
        character = ExcelCommentCharacter.parse(character_node)
        comment.characters << character.dup
      end
      comment
    end
  end
end
