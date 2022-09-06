# frozen_string_literal: true

require_relative 'presentation_comments/presentation_comment'
module OoxmlParser
  # Class for parsing `comment1.xml` file
  class PresentationComments < OOXMLDocumentObject
    # @return [Array<PresentationComment>] list of comments
    attr_reader :list

    def initialize(parent: nil)
      @list = []
      super
    end

    # Parse PresentationComments object
    # @param file [Nokogiri::XML:Element] node to parse
    # @return [PresentationComments] result of parsing
    def parse(file = "#{root_object.unpacked_folder}/#{root_object.root_subfolder}/comments/comment1.xml")
      return nil unless File.exist?(file)

      document = parse_xml(File.open(file))
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
