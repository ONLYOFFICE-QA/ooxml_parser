# Chart Label Properties
module OoxmlParser
  class DisplayLabelsProperties
    attr_accessor :position, :show_legend_key, :show_values, :show_x_axis_name, :show_y_axis_name

    def initialize(show_legend_key = false, show_values = false)
      @show_legend_key = show_legend_key
      @show_values = show_values
    end

    def self.parse(display_labels_node)
      display_labels_properties = DisplayLabelsProperties.new
      display_labels_node.xpath('*').each do |display_label_property_node|
        case display_label_property_node.name
        when 'dLblPos'
          display_labels_properties.position = Alignment.parse(display_label_property_node.attribute('val'))
        when 'showLegendKey'
          display_labels_properties.show_legend_key = true if display_label_property_node.attribute('val').value == '1'
        when 'showVal'
          display_labels_properties.show_values = true if display_label_property_node.attribute('val').value == '1'
        end
      end
      display_labels_properties
    end
  end
end
