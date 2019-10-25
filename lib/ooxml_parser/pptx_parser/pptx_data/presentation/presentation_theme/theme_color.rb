# frozen_string_literal: true

module OoxmlParser
  # Class for parsing ThemeColor tags
  class ThemeColor < OOXMLDocumentObject
    attr_accessor :type, :value, :color

    def initialize(type: '', color: nil, parent: nil)
      @type = type
      @color = color
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
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'sysClr'
          @type = :system
          @value = node_child.attribute('val').value
          @color = Color.new(parent: self).parse_hex_string(node_child.attribute('lastClr').value.to_s) unless node_child.attribute('lastClr').nil?
        when 'srgbClr'
          @type = :rgb
          @color = Color.new(parent: self).parse_hex_string(node_child.attribute('val').value.to_s)
        end
      end
      self
    end
  end
end
