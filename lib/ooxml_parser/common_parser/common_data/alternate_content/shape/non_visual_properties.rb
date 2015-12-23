require_relative 'shape_placeholder'
module OoxmlParser
  class NonVisualProperties < OOXMLDocumentObject
    attr_accessor :placeholder, :is_photo, :user_drawn

    def self.parse(non_visual_properties_node)
      non_visual_properties = NonVisualProperties.new
      non_visual_properties_node.xpath('*').each do |non_visual_properties_node_child|
        case non_visual_properties_node_child.name
        when 'ph'
          non_visual_properties.placeholder = ShapePlaceholder.parse(non_visual_properties_node_child)
        end
      end
      non_visual_properties.is_photo = OOXMLDocumentObject.option_enabled?(non_visual_properties_node, 'isPhoto')
      non_visual_properties.user_drawn = OOXMLDocumentObject.option_enabled?(non_visual_properties_node, 'userDrawn')
      non_visual_properties
    end
  end
end
