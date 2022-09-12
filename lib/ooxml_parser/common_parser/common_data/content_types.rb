# frozen_string_literal: true

require_relative 'content_types/content_type_default'
require_relative 'content_types/content_type_override'
module OoxmlParser
  # Class for data from `[Content_Types].xml` file
  class ContentTypes < OOXMLDocumentObject
    # @return [Array] list of content types
    attr_accessor :content_types_list

    def initialize(parent: nil)
      @content_types_list = []
      super
    end

    # @return [Array] accessor
    def [](key)
      @content_types_list[key]
    end

    # Parse ContentTypes object
    # @return [ContentTypes] result of parsing
    def parse
      doc = parse_xml("#{root_object.unpacked_folder}/[Content_Types].xml")
      node = doc.xpath('*').first

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'Default'
          @content_types_list << ContentTypeDefault.new(parent: self).parse(node_child)
        when 'Override'
          @content_types_list << ContentTypeOverride.new(parent: self).parse(node_child)
        end
      end
      self
    end

    # Get content definition by type
    # @param [String] type of definition
    # @return [Object] resulting objects
    def by_type(type)
      @content_types_list.select { |item| item.content_type == type }
    end
  end
end
