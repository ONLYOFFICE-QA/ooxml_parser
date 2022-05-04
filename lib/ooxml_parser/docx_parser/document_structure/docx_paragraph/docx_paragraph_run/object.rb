# frozen_string_literal: true

require_relative 'object/ole_object'
module OoxmlParser
  # Class for parsing `w:object`
  class RunObject < OOXMLDocumentObject
    # @return [OleObject] ole object of object
    attr_accessor :ole_object

    # Parse RunObject object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [RunObject] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'OLEObject'
          @ole_object = OleObject.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
