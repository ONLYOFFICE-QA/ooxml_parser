# Class for color transformations
module OoxmlParser
  class ColorProperties
    attr_accessor :alpha, :luminance_modulation, :luminance_offset
    attr_accessor :tint

    def self.parse(color_properties_node)
      properties = ColorProperties.new
      color_properties_node.xpath('*').each do |color_properties_node_child|
        case color_properties_node_child.name
        when 'alpha'
          properties.alpha = (color_properties_node_child.attribute('val').value.to_f / 1_000.0).round
        when 'lumMod'
          properties.luminance_modulation = color_properties_node_child.attribute('val').value.to_f / 100_000.0
        when 'lumOff'
          properties.luminance_offset = color_properties_node_child.attribute('val').value.to_f / 100_000.0
        when 'tint'
          properties.tint = color_properties_node_child.attribute('val').value.to_f / 100_000.0
        end
      end
      properties
    end
  end
end
