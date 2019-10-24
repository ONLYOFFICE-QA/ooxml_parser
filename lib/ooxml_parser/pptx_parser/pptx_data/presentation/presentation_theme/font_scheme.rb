# frozen_string_literal: true

require_relative 'font_scheme/font_collection'
module OoxmlParser
  # Class for parsing `fontScheme` tag
  class FontScheme < OOXMLDocumentObject
    # @return [String] name of font
    attr_reader :name
    # @return [FontCollection] major font data
    attr_reader :major_font
    # @return [FontCollection] minor font data
    attr_reader :minor_font

    # Parse FontScheme object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [FontScheme] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'name'
          @name = value.value.to_s
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'majorFont'
          @major_font = FontCollection.new(parent: self).parse(node_child)
        when 'minorFont'
          @minor_font = FontCollection.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
