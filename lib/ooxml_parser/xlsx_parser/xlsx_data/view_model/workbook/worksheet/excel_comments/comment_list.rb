# frozen_string_literal: true

module OoxmlParser
  # Single author `commentList`
  class CommentList < OOXMLDocumentObject
    attr_reader :comments

    def initialize(parent: nil)
      @comments = []
      super
    end

    # Parse Author object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [CommentList] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'comment'
          @comments << ExcelComment.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
