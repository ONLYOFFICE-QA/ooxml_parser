require_relative 'presentation_comments/presentation_comment'
module OoxmlParser
  # Class for parsing `comment1.xml` file
  class PresentationComments < OOXMLDocumentObject
    # @return [Array<PresentationComment>] list of comments
    attr_reader :list

    def initialize(parent: nil)
      @list = []
      @parent = parent
    end

    # Parse PresentationComments object
    # @param file [Nokogiri::XML:Element] node to parse
    # @return [PresentationComments] result of parsing
    def parse(file = "#{OOXMLDocumentObject.path_to_folder}/#{OOXMLDocumentObject.root_subfolder}/comments/comment1.xml")
      return nil unless File.exist?(file)

      document = Nokogiri::XML(File.open(file))
      node = document.xpath('*').first

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'cm'
          @list << PresentationComment.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
