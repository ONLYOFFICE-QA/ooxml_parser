# frozen_string_literal: true

module OoxmlParser
  # Stuff for help working with paragraph
  module DocxParagraphHelper
    # @return [Nil, CommentExtended] extended data for this comment
    def comment_extend_data
      return if @paragraph_id.nil?

      root_object.comments_extended.by_id(@paragraph_id)
    end
  end
end
