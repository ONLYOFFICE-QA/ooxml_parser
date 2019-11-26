# frozen_string_literal: true

require_relative 'old_docx_group_properties'
require_relative 'old_docx_group_element'
module OoxmlParser
  # Fallback DOCX group data
  class OldDocxGroup < OOXMLDocumentObject
    attr_accessor :elements, :properties

    def initialize(properties = OldDocxGroupProperties.new, parent: nil)
      @properties = properties
      @elements = []
      @parent = parent
    end

    # Parse OldDocxGroup object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [OldDocxGroup] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'shape'
          element = OldDocxGroupElement.new(:shape)
          element.object = OldDocxShape.new(parent: self).parse(node_child)
          @elements << element
        when 'wrap'
          @properties.wrap = node_child.attribute('type').value.to_sym unless node_child.attribute('type').nil?
        when 'group'
          @elements << OldDocxGroup.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
