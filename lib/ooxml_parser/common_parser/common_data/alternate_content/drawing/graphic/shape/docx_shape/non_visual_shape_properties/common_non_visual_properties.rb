module OoxmlParser
  class CommonNonVisualProperties
    attr_accessor :name, :id, :description, :on_click_hyperlink, :hyperlink_for_hover

    def initialize(id = '', name = '')
      @id = id
      @name = name
    end

    def self.parse(common_non_visual_properties_node)
      non_visual_properties = CommonNonVisualProperties.new
      non_visual_properties.name = common_non_visual_properties_node.attribute('name').value
      non_visual_properties.id = common_non_visual_properties_node.attribute('id').value
      common_non_visual_properties_node.xpath('*').each do |cnv_props_node_child|
        case cnv_props_node_child.name
        when 'hlinkClick'
          non_visual_properties.on_click_hyperlink = Hyperlink.parse(cnv_props_node_child)
        when 'hlinkHover'
          non_visual_properties.hyperlink_for_hover = HyperlinkForHover.parse(cnv_props_node_child)
        end
      end
      non_visual_properties
    end
  end
end
