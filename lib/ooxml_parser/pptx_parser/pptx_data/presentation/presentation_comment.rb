require_relative 'presentation_comment/presentation_comment_author'
module OoxmlParser
  class PresentationComment < OOXMLDocumentObject
    attr_accessor :author, :position, :text

    def initialize(author)
      @author = author
    end

    def self.parse(comment_node, comment_authors)
      comment = PresentationComment.new(comment_authors[comment_node.attribute('authorId').value])
      comment_node.xpath('*').each do |comment_node_child|
        case comment_node_child.name
        when 'pos'
          comment.position = OOXMLCoordinates.parse(comment_node_child, x_attr: 'x', y_attr: 'y')
        when 'text'
          comment.text = comment_node_child.text.to_s
        end
      end
      comment
    end

    def self.parse_list
      File.exist?(OOXMLDocumentObject.path_to_folder + 'ppt/comments/comment1.xml') ? doc = Nokogiri::XML(File.open(OOXMLDocumentObject.path_to_folder + '/ppt/comments/comment1.xml')) : (return [])
      comment_authors = PresentationCommentAuthor.parse
      comments = []
      doc.xpath('p:cmLst/p:cm').each { |comment_node| comments << PresentationComment.parse(comment_node, comment_authors) }
      comments
    end
  end
end
