require_relative 'non_visual_shape_properties/common_non_visual_properties'
require_relative 'non_visual_shape_properties/non_visual_properties'
module OoxmlParser
  class NonVisualShapeProperties
    attr_accessor :common_properties, :non_visual_properties

    def self.parse(nv_shape_props_node)
      non_visual_properties = NonVisualShapeProperties.new
      nv_shape_props_node.xpath('*').each do |nv_props_node_child|
        case nv_props_node_child.name
        when 'cNvPr'
          non_visual_properties.common_properties = CommonNonVisualProperties.parse(nv_props_node_child)
        when 'nvPr'
          non_visual_properties.non_visual_properties = NonVisualProperties.new(parent: non_visual_properties).parse(nv_props_node_child)
        end
      end
      non_visual_properties
    end
  end
end
