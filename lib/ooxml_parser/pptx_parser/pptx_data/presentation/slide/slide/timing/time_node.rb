require_relative 'time_node/common_timing'
require_relative 'animation_effect/animation_effect'
require_relative 'set_time_node/set_time_node'
module OoxmlParser
  class TimeNode
    attr_accessor :type, :common_time_node, :previous_conditions_list, :next_conditions_list

    def initialize(type = nil, common_time_node = nil, previous_conditions_list = [], next_conditions_list = [])
      @type = type
      @common_time_node = common_time_node
      @previous_conditions_list = previous_conditions_list
      @next_conditions_list = next_conditions_list
    end

    def self.parse(time_node, type)
      time = TimeNode.new
      time.type = type
      time_node.xpath('*').each do |time_node_child|
        case time_node_child.name
        when 'cTn'
          time.common_time_node = CommonTiming.new(parent: time).parse(time_node_child)
        when 'prevCondLst'
          time.previous_conditions_list = Condition.parse_list(time_node_child)
        end
      end
      time
    end

    def self.parse_list(timing_list_node)
      timings = []
      timing_list_node.xpath('*').each do |time_node|
        case time_node.name
        when 'par'
          timings << TimeNode.parse(time_node, :parallel)
        when 'seq'
          timings << TimeNode.parse(time_node, :sequence)
        when 'excl'
          timings << TimeNode.parse(time_node, :exclusive)
        when 'anim'
          timings << TimeNode.parse(time_node, :animate)
        when 'set'
          timings << SetTimeNode.parse(time_node)
        when 'animEffect'
          timings << AnimationEffect.parse(time_node)
        when 'video'
          timings << :video
        when 'audio'
          timings << :audio
        end
      end
      timings
    end
  end
end
