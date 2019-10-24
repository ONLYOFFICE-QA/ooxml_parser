# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `w:commentRangeEnd` tags
  class CommentRangeEnd < OOXMLDocumentObject
    # @return [Integer] id of bookmark
    attr_reader :id

    # Parse CommentRangeEnd object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [BookmarkStart] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'id'
          @id = value.value.to_i
        end
      end
      self
    end
  end
end
