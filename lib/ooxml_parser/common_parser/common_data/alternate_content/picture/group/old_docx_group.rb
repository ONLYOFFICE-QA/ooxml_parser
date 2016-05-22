require_relative 'old_docx_group_properties'
require_relative 'old_docx_group_element'
# Fallback DOCX group data
module OoxmlParser
  class OldDocxGroup
    attr_accessor :elements, :properties

    def initialize(properties = OldDocxGroupProperties.new, elements = [])
      @properties = properties
      @elements = elements
    end

    def self.parse(group_node)
      group = OldDocxGroup.new
      group_node.xpath('*').each do |group_node_child|
        case group_node_child.name
        when 'shape'
          element = OldDocxGroupElement.new(:shape)
          element.object = OldDocxShape.parse(group_node_child)
          group.elements << element
        when 'wrap'
          group.properties.wrap = group_node_child.attribute('type').value.to_sym unless group_node_child.attribute('type').nil?
        when 'group'
          element = parse(group_node_child)
          group.elements << element
        end
      end
      group
    end
  end
end
