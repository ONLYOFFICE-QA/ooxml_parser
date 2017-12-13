module OoxmlParser
  # Class for parsing ThemeColor tags
  class ThemeColor < OOXMLDocumentObject
    attr_accessor :type, :value, :color

    def initialize(parent: nil)
      @type = ''
      @parent = parent
    end

    def ==(other)
      if other.is_a?(Color)
        @color == other
      else
        all_instance_variables = instance_variables
        all_instance_variables.each do |current_attributes|
          return false unless instance_variable_get(current_attributes) == other.instance_variable_get(current_attributes)
        end
        true
      end
    end

    # Parse ThemeColor
    # @param node [Nokogiri::XML::Element] node to parse
    # @return [ThemeColor] result of parsing
    def parse(color_node)
      color_node.xpath('*').each do |color_node_child|
        case color_node_child.name
        when 'sysClr'
          @type = :system
          @value = color_node_child.attribute('val').value
          @color = Color.new(parent: self).parse_hex_string(color_node_child.attribute('lastClr').value.to_s) unless color_node_child.attribute('lastClr').nil?
        when 'srgbClr'
          @type = :rgb
          @color = Color.new(parent: self).parse_hex_string(color_node_child.attribute('val').value.to_s)
        end
      end
      self
    end
  end
end
