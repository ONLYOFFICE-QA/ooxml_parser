# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `w:CommentRangeStart` tags
  class CommentRangeStart < OOXMLDocumentObject
    # @return [Integer] id of bookmark
    attr_reader :id

    # Parse CommentRangeStart object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [CommentRangeStart] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'id'
          @id = value.value.to_i
        end
      end
      self
    end

    # @return [Comment] object of current comment range
    def comment
      root_object.comments.comments_array.detect { |comment| comment.id == id }
    end
  end
end
