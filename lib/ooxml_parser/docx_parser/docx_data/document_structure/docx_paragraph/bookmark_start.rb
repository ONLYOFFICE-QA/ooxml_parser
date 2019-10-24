# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `w:bookmarkStart` tags
  class BookmarkStart < OOXMLDocumentObject
    # @return [Integer] id of bookmark
    attr_reader :id
    # @return [String] name of bookmark
    attr_reader :name

    # Parse BookmarkStart object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [BookmarkStart] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'id'
          @id = value.value.to_i
        when 'name'
          @name = value.value
        end
      end
      self
    end
  end
end
