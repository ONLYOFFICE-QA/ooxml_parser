# Comment Data
module OoxmlParser
  class Comment < OOXMLDocumentObject
    attr_accessor :id, :paragraphs

    def initialize(id = nil, paragraphs = [])
      @id = id
      @paragraphs = paragraphs
    end

    def self.parse_list
      comments = []
      comments_filename = "#{OOXMLDocumentObject.path_to_folder}/#{OOXMLDocumentObject.root_subfolder}/comments.xml"
      return [] unless File.exist?(comments_filename)
      doc = Nokogiri::XML(File.open(comments_filename))
      doc.search('//w:comments').each do |document|
        document.xpath('w:comment').each do |comment_tag|
          comment = Comment.new(comment_tag.attribute('id').value)
          comment_tag.xpath('w:p').each_with_index do |p, index|
            comment.paragraphs << DocxParagraph.new.parse(p, index)
          end
          comments << comment.dup
        end
      end
      comments
    end
  end
end
