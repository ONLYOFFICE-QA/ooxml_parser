# frozen_string_literal: true

module OoxmlParser
  # Comment Data
  class Comment < OOXMLDocumentObject
    # @return [String] name of the author
    attr_reader :author
    # @return [String] date of the comment, in string, not parsed
    attr_reader :date_string
    # @return [Integer] id of comment
    attr_reader :id
    # @return [String] initials of the author
    attr_reader :initials
    # @return [Array<DocxParagraph>] array of paragraphs
    attr_reader :paragraphs

    def initialize(id = nil, paragraphs = [], parent: nil)
      @id = id
      @paragraphs = paragraphs
      super(parent: parent)
    end

    # Parse Comment object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Comment] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'author'
          @author = value.value.to_s
        when 'date'
          @date_string = value.value.to_s
        when 'id'
          @id = value.value.to_i
        when 'initials'
          @initials = value.value.to_s
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'p'
          @paragraphs << DocxParagraph.new.parse(node_child, 0, parent: self)
        end
      end
      self
    end
  end
end
