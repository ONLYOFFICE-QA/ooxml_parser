# frozen_string_literal: true

module OoxmlParser
  # Comment Data
  class Comment < OOXMLDocumentObject
    # @return [Integer] id of comment
    attr_reader :id
    # @return [Array<DocxParagraph>] array of paragraphs
    attr_reader :paragraphs

    def initialize(id = nil, paragraphs = [], parent: nil)
      @id = id
      @paragraphs = paragraphs
      @parent = parent
    end

    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'id'
          @id = value.value.to_i
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
