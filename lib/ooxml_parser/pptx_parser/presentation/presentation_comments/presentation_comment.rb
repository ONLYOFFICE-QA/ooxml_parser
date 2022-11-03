# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `p:cm` tag
  class PresentationComment < OOXMLDocumentObject
    # @return [Integer] id of author
    attr_reader :author_id
    # @return [OOXMLCoordinates] position of comment
    attr_reader :position
    # @return [String] text of comment
    attr_reader :text

    # @return [CommentAuthor] author of comment
    def author
      root_object.comment_authors.author_by_id(@author_id)
    end

    # Parse PresentationComment object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [PresentationComment] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'authorId'
          @author_id = value.value.to_i
        end
      end
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'pos'
          @position = OOXMLCoordinates.new(parent: self).parse(node_child)
        when 'text'
          @text = node_child.text.to_s
        end
      end
      self
    end
  end
end
