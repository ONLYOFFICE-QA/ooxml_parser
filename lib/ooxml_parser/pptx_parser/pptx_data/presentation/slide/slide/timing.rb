require_relative 'timing/time_node'
module OoxmlParser
  class Timing
    attr_accessor :time_node_list, :build_list, :extension_list

    def initialize(time_node_list = [], build_list = [], extension_list = [])
      @time_node_list = time_node_list
      @build_list = build_list
      @extension_list = extension_list
    end

    def self.parse(timing_node)
      timing = Timing.new
      timing_node.xpath('*').each do |timing_node_child|
        case timing_node_child.name
        when 'tnLst'
          timing.time_node_list = TimeNode.parse_list(timing_node_child)
        end
      end
      timing
    end
  end
end
