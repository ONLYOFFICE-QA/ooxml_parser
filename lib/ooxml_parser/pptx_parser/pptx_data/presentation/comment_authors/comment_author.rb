# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `cmAuthor` tag
  class CommentAuthor < OOXMLDocumentObject
    # @return [Integer] Comment Author ID
    attr_reader :id
    # @return [String] Author Name
    attr_reader :name
    # @return [String] Author Initials
    attr_reader :initials
    # @return [Integer] Index of Comment Author's last comment
    attr_reader :last_index
    # @return [Integer] Comment Author Color Index
    attr_reader :color_index

    # Parse CommentAuthor object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [CommentAuthor] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'id'
          @id = value.value.to_i
        when 'name'
          @name = value.value.to_s
        when 'initials'
          @initials = value.value.to_s
        when 'lastIdx'
          @last_index = value.value.to_i
        when 'clrIdx'
          @color_index = value.value.to_i
        end
      end
      self
    end
  end
end
