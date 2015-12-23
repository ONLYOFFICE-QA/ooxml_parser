module OoxmlParser
  class GroupingProperties
    attr_accessor :transform

    def self.parse(grouping_properties_node)
      grouping_properties = GroupingProperties.new
      grouping_properties_node.xpath('*').each do |grouping_property_node|
        case grouping_property_node.name
        when 'xfrm'
          grouping_properties.transform = TransformEffect.parse(grouping_property_node)
        end
      end
      grouping_properties
    end
  end
end
