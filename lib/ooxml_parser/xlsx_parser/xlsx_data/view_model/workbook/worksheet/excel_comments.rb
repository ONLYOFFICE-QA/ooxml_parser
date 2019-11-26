# frozen_string_literal: true

require_relative 'excel_comments/author'
require_relative 'excel_comments/comment_list'
require_relative 'excel_comments/excel_comment'
module OoxmlParser
  # All Comments of single XLSX
  class ExcelComments < OOXMLDocumentObject
    attr_accessor :authors
    # @return [CommentList] list of comments
    attr_reader :comment_list

    def initialize(parent: nil)
      @authors = []
      @comment_list = []
      @parent = parent
    end

    def comments
      comment_list.comments
    end

    extend Gem::Deprecate
    deprecate :comments, 'comment_list.comments', 2020, 1

    def parse(file)
      doc = parse_xml(file)
      node = doc.xpath('*').first
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'authors'
          @authors << Author.new(parent: self).parse(node_child)
        when 'commentList'
          @comment_list = CommentList.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
