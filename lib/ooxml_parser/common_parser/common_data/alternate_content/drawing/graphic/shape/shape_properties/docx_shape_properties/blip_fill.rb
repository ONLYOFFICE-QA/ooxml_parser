# frozen_string_literal: true

require_relative 'blip_fill/blip'
module OoxmlParser
  # Class for parsing `a:blipFill` tag
  class BlipFill < OOXMLDocumentObject
    attr_accessor :blip
    # @return [Stretch]
    attr_accessor :stretch

    # Parse BlipFill object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [BlipFill] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'blip'
          @blip = Blip.new(parent: self).parse(node_child)
        when 'stretch'
          @stretch = Stretch.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
