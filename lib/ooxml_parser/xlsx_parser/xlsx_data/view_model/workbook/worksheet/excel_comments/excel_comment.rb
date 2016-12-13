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
        comment.characters << ParagraphRun.parse(character_node)
      end
      comment
    end
  end
end
