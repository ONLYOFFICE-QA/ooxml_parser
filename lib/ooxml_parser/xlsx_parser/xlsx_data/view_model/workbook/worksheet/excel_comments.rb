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
      path_to_comments_xml = ''
      return nil unless File.exist?(path_to_folder + "xl/worksheets/_rels/#{file_name}.rels")
      relationships = XmlSimple.xml_in(File.open(path_to_folder + "xl/worksheets/_rels/#{file_name}.rels"))
      if relationships['Relationship']
        relationships['Relationship'].each do |relationship|
          path_to_comments_xml = path_to_folder + 'xl/' + relationship['Target'].gsub('..', '') if File.basename(relationship['Target']).include?('comment')
        end
      end
      ExcelComments.parse(path_to_comments_xml) unless path_to_comments_xml == ''
    end
  end
end
