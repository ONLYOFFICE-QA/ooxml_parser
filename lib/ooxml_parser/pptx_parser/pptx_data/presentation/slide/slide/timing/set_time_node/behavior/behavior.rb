# frozen_string_literal: true

require_relative 'target_element'
module OoxmlParser
  # Class for data for Behavior
  class Behavior < OOXMLDocumentObject
    attr_accessor :common_time_node, :attribute_name_list, :target

    def initialize(parent: nil)
      @attribute_name_list = []
      @parent = parent
    end

    # Parse Behavior object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Behavior] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'cTn'
          @common_time_node = CommonTiming.new(parent: self).parse(node_child)
        when 'tgtEl'
          @target = TargetElement.new(parent: self).parse(node_child)
        when 'attrNameLst'
          node_child.xpath('p:attrName').each do |attribute_name_node|
            @attribute_name_list << attribute_name_node.text
          end
        end
      end
      self
    end
  end
end
