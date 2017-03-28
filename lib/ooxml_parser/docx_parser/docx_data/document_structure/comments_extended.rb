require_relative 'comments_extended/comment_extended'
module OoxmlParser
  # Class for parsing `commentsExtended.xml` file
  class CommentsExtended < OOXMLDocumentObject
    def initialize(parent: nil)
      @comments_extended_array = []
      @parent = parent
    end

    # @return [Array, CommentsExtended] accessor
    def [](key)
      @comments_extended_array[key]
    end

    # Parse CommentsExtended object
    # @return [CommentsExtended] result of parsing
    def parse
      file_to_parse = OOXMLDocumentObject.path_to_folder + 'word/commentsExtended.xml'
      return nil unless File.exist?(file_to_parse)
      doc = Nokogiri::XML(File.open(file_to_parse))
      doc.xpath('w15:commentsEx/*').each do |node_child|
        case node_child.name
        when 'commentEx'
          @comments_extended_array << CommentExtended.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
