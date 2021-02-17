# frozen_string_literal: true

module OoxmlParser
  # Stuff for help working with paragraph
  module DocxParagraphHelper
    # @return [Nil, CommentExtended] extended data for this comment
    def comment_extend_data
      return if @paragraph_id.nil?

      root_object.comments_extended.by_id(@paragraph_id)
    end

    # Temp method to return background color
    # Need to be compatible with older versions
    # @return [OoxmlParser::Color]
    def background_color
      paragraph_properties.shade.to_background_color
    end
  end
end
