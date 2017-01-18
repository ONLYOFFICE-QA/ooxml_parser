require_relative 'excel_comments/excel_comment'
# All Comments of single XLSX
module OoxmlParser
  class ExcelComments
    attr_accessor :authors, :comments

    def initialize(authors = [], comments = [])
      @authors = authors
      @comments = comments
    end

    def self.parse(file)
      doc = Nokogiri::XML(File.open(file))
      excel_comments = ExcelComments.new
      comments_node = doc.search('//xmlns:comments').first
      comments_node.xpath('*').each do |comments_node_child|
        case comments_node_child.name
        when 'authors'
          comments_node_child.xpath('xmlns:author').each do |author_node|
            excel_comments.authors << author_node.text
          end
        when 'commentList'
          comments_node_child.xpath('xmlns:comment').each do |comment_node|
            excel_comments.comments << ExcelComment.new(parent: excel_comments).parse(comment_node)
          end
        end
      end
      excel_comments
    end

    def self.parse_file(file_name, path_to_folder)
      file = path_to_folder + "xl/worksheets/_rels/#{file_name}.rels"
      return nil unless File.exist?(file)
      relationships = Relationships.parse_rels(file)
      target = relationships.target_by_type('comment')
      return if target.nil?
      comment_file = "#{OOXMLDocumentObject.path_to_folder}/#{OOXMLDocumentObject.root_subfolder}/#{target.gsub('..', '')}"
      ExcelComments.parse(comment_file)
    end
  end
end
