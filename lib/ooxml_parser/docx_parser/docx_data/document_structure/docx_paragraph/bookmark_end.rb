# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `w:bookmarkEnd` tags
  class BookmarkEnd < OOXMLDocumentObject
    # @return [Integer] id of bookmark
    attr_reader :id

    # Parse BookmarkStart object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [BookmarkEnd] result of parsing
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
