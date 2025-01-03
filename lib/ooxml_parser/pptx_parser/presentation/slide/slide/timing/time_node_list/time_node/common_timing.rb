# frozen_string_literal: true

require_relative 'common_timing/condition_list'
module OoxmlParser
  # Class for data of CommonTiming
  class CommonTiming < OOXMLDocumentObject
    attr_accessor :id, :duration, :restart, :children, :start_conditions, :end_conditions

    def initialize(parent: nil)
      @children = []
      @start_conditions = []
      @end_conditions = []
      super
    end

    # Parse CommonTiming object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [CommonTiming] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'dur'
          @duration = value.value
        when 'restart'
          @restart = value.value
        when 'id'
          @id = value.value
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'stCondLst'
          @start_conditions = ConditionList.new(parent: self).parse(node_child)
        when 'endCondLst'
          @end_conditions = ConditionList.new(parent: self).parse(node_child)
        when 'childTnLst'
          @children = TimeNodeList.new(parent: self).parse(node_child).elements
        end
      end
      self
    end
  end
end
