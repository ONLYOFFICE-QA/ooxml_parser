# frozen_string_literal: true

require_relative 'comment_authors/comment_author'
module OoxmlParser
  # Class for parsing `commentAuthors.xml` file
  class CommentAuthors < OOXMLDocumentObject
    # @return [Array<CommentAuthor>] list of comment authors
    attr_reader :list

    def initialize(parent: nil)
      @list = []
      super
    end

    # Parse CommentAuthors object
    # @param file [Nokogiri::XML:Element] node to parse
    # @return [CommentAuthors] result of parsing
    def parse(file = "#{root_object.unpacked_folder}/#{root_object.root_subfolder}/commentAuthors.xml")
      return nil unless File.exist?(file)

      document = parse_xml(file)
      node = document.xpath('*').first

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'cmAuthor'
          @list << CommentAuthor.new(parent: self).parse(node_child)
        end
      end
      self
    end

    # @param id [Integer] id of author
    # @return [CommentAuthor] author by id
    def author_by_id(id)
      list.detect { |author| author.id == id }
    end
  end
end
