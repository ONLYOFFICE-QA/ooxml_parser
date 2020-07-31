# frozen_string_literal: true

module OoxmlParser
  # Module with Accessors to DocxParagraph
  module DocxParagraphAccessors
    # @return [Array<OOXMLDocumentObject>] array of child objects that contains data
    def nonempty_runs
      @character_style_array.select do |cur_run|
        case cur_run
        when DocxParagraphRun, ParagraphRun
          !cur_run.empty?
        when DocxFormula, StructuredDocumentTag, BookmarkStart, BookmarkEnd, CommentRangeStart, CommentRangeEnd
          true
        end
      end
    end

    extend Gem::Deprecate
    deprecate :page_numbering, 'field_simple.page_numbering?', 2020, 1

    # @return [OoxmlParser::StructuredDocumentTag] Return first sdt element
    def sdt
      @character_style_array.each do |cur_element|
        return cur_element if cur_element.is_a?(StructuredDocumentTag)
      end
      nil
    end
    deprecate :sdt, 'nonempty_runs[i]', 2020, 1

    # @return [OoxmlParser::FrameProperties] Return frame properties
    def frame_properties
      paragraph_properties.frame_properties
    end
    deprecate :frame_properties, 'paragraph_properties.frame_properties', 2020, 1
  end
end
