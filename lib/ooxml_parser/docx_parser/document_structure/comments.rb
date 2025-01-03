# frozen_string_literal: true

require_relative 'comments/comment'
module OoxmlParser
  # Class for parsing `comments.xml` file
  class Comments < OOXMLDocumentObject
    # @return [Array<Comment>] list of comments
    attr_reader :comments_array

    def initialize(params = {})
      @comments_array = []
      super(parent: params[:parent])
      @file = params.fetch(:file, "#{root_object.unpacked_folder}word/comments.xml")
    end

    # @return [Comment] accessor
    def [](key)
      @comments_array[key]
    end

    # Parse CommentsExtended object
    # @return [Comments] result of parsing
    def parse
      return nil unless File.file?(@file)

      doc = parse_xml(@file)
      doc.xpath('w:comments/*').each do |node_child|
        case node_child.name
        when 'comment'
          @comments_array << Comment.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
