# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <buBlip> tag
  class BulletImage < OOXMLDocumentObject
    # @return [FileReference] image structure
    attr_reader :file_reference

    # Parse BulletImage object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [BulletImage] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'blip'
          @file_reference = FileReference.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
