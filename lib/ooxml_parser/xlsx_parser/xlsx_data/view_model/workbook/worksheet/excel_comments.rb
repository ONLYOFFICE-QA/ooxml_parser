require_relative 'excel_comments/author'
require_relative 'excel_comments/comment_list'
require_relative 'excel_comments/excel_comment'
# All Comments of single XLSX
module OoxmlParser
  class ExcelComments
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
      doc = Nokogiri::XML(File.open(file))
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

    def self.parse_file(file_name, path_to_folder)
      file = path_to_folder + "xl/worksheets/_rels/#{file_name}.rels"
      return nil unless File.exist?(file)
      relationships = Relationships.parse_rels(file)
      target = relationships.target_by_type('comment')
      return if target.nil?
      comment_file = "#{OOXMLDocumentObject.path_to_folder}/#{OOXMLDocumentObject.root_subfolder}/#{target.gsub('..', '')}"
      ExcelComments.new.parse(comment_file)
    end
  end
end
