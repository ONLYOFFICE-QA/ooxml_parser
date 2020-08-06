# frozen_string_literal: true

require_relative 'time_node/common_timing'
require_relative 'animation_effect/animation_effect'
require_relative 'set_time_node/set_time_node'
module OoxmlParser
  # Class for parsing TimeNode tags
  class TimeNode < OOXMLDocumentObject
    attr_accessor :type, :common_time_node, :previous_conditions_list, :next_conditions_list

    def initialize(type = nil, parent: nil)
      @type = type
      @previous_conditions_list = []
      @next_conditions_list = []
      super(parent: parent)
    end

    # Parse TimeNode
    # @param node [Nokogiri::XML::Element] node to parse
    # @return [TimeNode] value of SheetFormatProperties
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'cTn'
          @common_time_node = CommonTiming.new(parent: self).parse(node_child)
        when 'prevCondLst'
          @previous_conditions_list = ConditionList.new(parent: self).parse(node_child)
        end
      end
      self
    end

    # Parse list of timing nodes
    # @param timing_list_node [Nokogiri::XML::Element] node to parse
    # @return [Array<TimeNode>] list of nodes
    def self.parse_list(timing_list_node)
      timings = []
      timing_list_node.xpath('*').each do |time_node|
        case time_node.name
        when 'par'
          timings << TimeNode.new(:parallel, parent: timings).parse(time_node)
        when 'seq'
          timings << TimeNode.new(:sequence, parent: timings).parse(time_node)
        when 'anim'
          timings << TimeNode.new(:animate, parent: timings).parse(time_node)
        when 'set'
          timings << SetTimeNode.new(parent: timings).parse(time_node)
        when 'animEffect'
          timings << AnimationEffect.new(parent: timings).parse(time_node)
        when 'video'
          timings << :video
        when 'audio'
          timings << TimeNode.new(:audio, parent: timings).parse(time_node)
        end
      end
      timings
    end
  end
end
