# frozen_string_literal: true

require_relative 'background_properties/stretch'
module OoxmlParser
  # Class for parsing `bgPr` tag
  class BackgroundProperties < OOXMLDocumentObject
    # @return [BlipFill]
    attr_accessor :blip_fill

    # Parse BackgroundProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [BackgroundProperties] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'blipFill'
          @blip_fill = BlipFill.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
