# Chart Label Properties
module OoxmlParser
  class DisplayLabelsProperties < OOXMLDocumentObject
    attr_accessor :position, :show_legend_key, :show_values, :show_x_axis_name, :show_y_axis_name
    # @return [True, False] is category name shown
    attr_accessor :show_category_name
    # @return [True, False] is series name shown
    attr_accessor :show_series_name

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
        when 'showCatName'
          display_labels_properties.show_category_name = option_enabled?(display_label_property_node)
        when 'showSerName'
          display_labels_properties.show_series_name = option_enabled?(display_label_property_node)
        end
      end
      display_labels_properties
    end
  end
end
