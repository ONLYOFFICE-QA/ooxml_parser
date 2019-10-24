# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `wp:docP` tags
  class DocProperties < OOXMLDocumentObject
    # @return [String] title property
    attr_accessor :title
    # @return [String] description property
    attr_accessor :description

    # Parse DocProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocProperties] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'title'
          @title = value.value
        when 'descr'
          @description = value.value
        end
      end
      self
    end
  end
end
