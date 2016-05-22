require_relative 'shape/shape_properties'
module OoxmlParser
  class Shape < OOXMLDocumentObject
    attr_accessor :type, :properties, :elements

    def initialize(type = nil, properties = ShapeProperties.new, elements = [])
      @type = type
      @properties = properties
      @elements = elements
    end

    def self.parse(shape_node, type, path)
      shape = Shape.new
      shape.type = type
      shape_node.attribute('style').value.to_s.split(';').each do |property|
        case
        when property.include?('margin-top')
          shape.properties.margins.top = property.split(':').last
        when property.include?('margin-left')
          shape.properties.margins.left = property.split(':').last
        when property.include?('margin-right')
          shape.properties.margins.right = property.split(':').last
        when property.include?('width')
          shape.properties.size.width = property.split(':').last
        when property.include?('height')
          shape.properties.size.height = property.split(':').last
        when property.include?('z-index')
          shape.properties.z_index = property.split(':').last
        when property.include?('position')
          shape.properties.position = property.split(':').last
        end
      end
      shape.properties.fill_color = Color.from_int16(shape_node.attribute('fillcolor').value.to_s.sub('#', '').split(' ').first) unless shape_node.attribute('fillcolor').nil?
      shape.properties.stroke.weight = shape_node.attribute('strokeweight').value unless shape_node.attribute('strokeweight').nil?
      shape.properties.stroke.color = Color.from_int16(shape_node.attribute('strokecolor').value.to_s.sub('#', '').split(' ').first) unless shape_node.attribute('strokecolor').nil?
      shape.elements = TextBox.parse_list(shape_node.xpath('v:textbox').first) unless shape_node.xpath('v:textbox').first.nil?
      shape
    end
  end
end
