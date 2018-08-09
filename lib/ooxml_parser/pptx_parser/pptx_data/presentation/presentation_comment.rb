require_relative 'presentation_comment/presentation_comment_author'
module OoxmlParser
  class PresentationComment < OOXMLDocumentObject
    attr_accessor :position, :text
    # @return [Integer] id of author
    attr_accessor :author_id

    # @return [CommentAuthor] author of comment
    def author
      root_object.comment_authors.author_by_id(@author_id)
    end

    def self.parse(comment_node, parent)
      comment = PresentationComment.new(parent: parent)
      comment_node.attributes.each do |key, value|
        case key
        when 'authorId'
          comment.author_id = value.value.to_i
        end
      end
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

    def self.parse_list(parent)
      File.exist?(OOXMLDocumentObject.path_to_folder + 'ppt/comments/comment1.xml') ? doc = Nokogiri::XML(File.open(OOXMLDocumentObject.path_to_folder + '/ppt/comments/comment1.xml')) : (return [])
      comments = []
      doc.xpath('p:cmLst/p:cm').each { |comment_node| comments << PresentationComment.parse(comment_node, parent) }
      comments
    end
  end
end
