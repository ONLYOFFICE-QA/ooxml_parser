module OoxmlParser
  class PresentationCommentAuthor < OOXMLDocumentObject
    attr_accessor :id, :name, :initials

    def initialize(id, name, initials)
      @id = id
      @name = name
      @initials = initials
    end

    def self.parse
      comment_authors = {}
      begin
        doc = Nokogiri::XML(File.open(OOXMLDocumentObject.path_to_folder + 'ppt/commentAuthors.xml'))
      rescue StandardError
        raise "Can't find commentAuthors.xml in #{OOXMLDocumentObject.path_to_folder}/ppt folder"
      end
      doc.xpath('p:cmAuthorLst/p:cmAuthor').each do |author_node|
        comment_authors[author_node.attribute('id').value] = PresentationCommentAuthor.new(
          author_node.attribute('id').value, author_node.attribute('name').value, author_node.attribute('initials').value)
      end
      comment_authors
    end
  end
end
