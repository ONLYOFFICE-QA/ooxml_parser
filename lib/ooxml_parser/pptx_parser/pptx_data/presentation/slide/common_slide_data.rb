# frozen_string_literal: true

require_relative 'common_slide_data/shape_tree'

module OoxmlParser
  # Class for parsing `cSld` tag
  class CommonSlideData < OOXMLDocumentObject
    # @return [ShapeTree] tree of shape
    attr_reader :shape_tree
    # @return [Background] background
    attr_reader :background

    # Parse CommonSlideData object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [CommonSlideData] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'spTree'
          @shape_tree = ShapeTree.new(parent: self).parse(node_child)
        when 'bg'
          @background = Background.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
