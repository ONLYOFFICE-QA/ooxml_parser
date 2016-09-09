module OoxmlParser
  class ThemeColor
    attr_accessor :type, :value, :color

    def initialize(type = '', color = nil)
      @type = type
      @color = color
    end

    def ==(other)
      if other.is_a?(Color)
        @color == other
      else
        all_instance_variables = instance_variables
        all_instance_variables.each do |current_attributes|
          unless instance_variable_get(current_attributes) == other.instance_variable_get(current_attributes)
            return false
          end
        end
        true
      end
    end

    def self.parse(color_node)
      theme_color = ThemeColor.new
      color_node.xpath('*').each do |color_node_child|
        case color_node_child.name
        when 'sysClr'
          theme_color.type = :system
          theme_color.value = color_node_child.attribute('val').value
          theme_color.color = Color.new(parent: theme_color).parse_hex_string(color_node_child.attribute('lastClr').value.to_s) unless color_node_child.attribute('lastClr').nil?
        when 'srgbClr'
          theme_color.type = :rgb
          theme_color.color = Color.new(parent: theme_color).parse_hex_string(color_node_child.attribute('val').value.to_s)
        end
      end
      theme_color
    end
  end
end
