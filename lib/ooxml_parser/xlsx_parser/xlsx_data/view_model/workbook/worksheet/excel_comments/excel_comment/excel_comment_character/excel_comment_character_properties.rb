# Properties of XLSX comment Character (Run)
module OoxmlParser
  class ExcelCommentCharacterProperties < OOXMLDocumentObject
    attr_accessor :size, :color, :font

    def initialize(size = '', color = nil, font = '')
      @size = size
      @color = color
      @font = font
    end

    def self.parse(properties_node)
      character_properties = ExcelCommentCharacterProperties.new
      properties_node.xpath('*').each do |properties_node_child|
        case properties_node_child.name
        when 'sz'
          character_properties.size = properties_node_child.attribute('val').value
        when 'color'
          character_properties.color = Color.parse_color_tag(properties_node_child)
        when 'rFont'
          character_properties.font = properties_node_child.attribute('val').value
        end
      end
      character_properties
    end
  end
end
