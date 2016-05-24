require_relative 'target_element'
module OoxmlParser
  class Behavior
    attr_accessor :common_time_node, :attribute_name_list, :target

    def initialize(common_time_node = nil, attribute_name_list = [], target = nil)
      @common_time_node = common_time_node
      @attribute_name_list = attribute_name_list
      @target = target
    end

    def self.parse(behavior_node)
      behavior = Behavior.new
      behavior_node.xpath('*').each do |behavior_node_child|
        case behavior_node_child.name
        when 'cTn'
          behavior.common_time_node = CommonTiming.parse(behavior_node_child)
        when 'tgtEl'
          behavior.target = TargetElement.parse(behavior_node_child)
        when 'attrNameLst'
          behavior_node_child.xpath('p:attrName').each do |attribute_name_node|
            behavior.attribute_name_list << attribute_name_node.text
          end
        end
      end
      behavior
    end
  end
end
