require_relative 'common_timing/condition'
module OoxmlParser
  class CommonTiming
    attr_accessor :id, :duration, :restart, :children, :start_conditions, :end_conditions

    def initialize
      @id = nil
      @duration = nil
      @restart = nil
      @children = []
      @start_conditions = []
      @end_conditions = []
    end

    def self.parse(common_time_node)
      common_timing = CommonTiming.new
      common_timing.duration = common_time_node.attribute('dur').value unless common_time_node.attribute('dur').nil?
      common_timing.restart = common_time_node.attribute('restart').value unless common_time_node.attribute('restart').nil?
      common_timing.id = common_time_node.attribute('id').value
      common_time_node.xpath('*').each do |common_time_node_child|
        case common_time_node_child.name
        when 'stCondLst'
          common_timing.start_conditions = Condition.parse_list(common_time_node_child)
        when 'endCondLst'
          common_timing.end_conditions = Condition.parse_list(common_time_node_child)
        when 'childTnLst'
          common_timing.children = TimeNode.parse_list(common_time_node_child)
        end
      end
      common_timing
    end
  end
end
