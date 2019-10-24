# frozen_string_literal: true

require_relative 'font_collection/text_font'
module OoxmlParser
  # Class for parsing `a:majorFont`, `a:minorFont` tag
  class FontCollection < OOXMLDocumentObject
    # @return [TextFont] latin text font
    attr_reader :latin
    # @return [TextFont] east asian text font
    attr_reader :east_asian
    # @return [TextFont] complex script text font
    attr_reader :complex_script

    # Parse FontCollection object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [FontCollection] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'latin'
          @latin = TextFont.new(parent: self).parse(node_child)
        when 'ea'
          @east_asian = TextFont.new(parent: self).parse(node_child)
        when 'cs'
          @complex_script = TextFont.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
