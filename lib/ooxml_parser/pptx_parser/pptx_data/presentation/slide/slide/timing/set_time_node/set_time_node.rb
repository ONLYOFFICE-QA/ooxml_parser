require_relative 'behavior/behavior'

module OoxmlParser
  class SetTimeNode
    attr_accessor :behavior, :to

    def initialize(behavior = nil)
      @behavior = behavior
    end

    def self.parse(set_time_node)
      set_time = SetTimeNode.new
      set_time_node.xpath('*').each do |set_time_node_child|
        case set_time_node_child.name
        when 'cBhvr'
          set_time.behavior = Behavior.parse(set_time_node_child)
        end
      end
      set_time
    end
  end
end
