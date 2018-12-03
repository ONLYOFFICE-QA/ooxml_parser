require_relative 'comments/comment'
module OoxmlParser
  # Class for parsing `comments.xml` file
  class Comments < OOXMLDocumentObject
    # @return [Array<Comment>] list of comments
    attr_reader :comments_array

    def initialize(parent: nil)
      @comments_array = []
      @parent = parent
    end

    # @return [Array, Comments] accessor
    def [](key)
      @comments_array[key]
    end

    # Parse CommentsExtended object
    # @return [Comments] result of parsing
    def parse
      file_to_parse = OOXMLDocumentObject.path_to_folder + 'word/comments.xml'
      return nil unless File.exist?(file_to_parse)

      doc = parse_xml(file_to_parse)
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
