# frozen_string_literal: true

require_relative 'background/background_properties'
module OoxmlParser
  # Class for parsing `bg` tag
  class Background < OOXMLDocumentObject
    attr_accessor :fill, :shade_to_title
    # @return [BackgroundProperties] properties
    attr_accessor :properties

    def initialize(type = nil, parent: nil)
      @type = type
      super(parent: parent)
    end

    # Parse Background object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Background] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'bgPr'
          @properties = BackgroundProperties.new(parent: self).parse(node_child)
        end
      end

      background_properties_node = node.xpath('p:bgPr').first
      if background_properties_node
        @shade_to_title = attribute_enabled?(background_properties_node, 'shadeToTitle')
        @fill = PresentationFill.new(parent: self).parse(background_properties_node)
      end
      self
    end
  end
end
