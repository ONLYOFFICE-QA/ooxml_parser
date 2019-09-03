require_relative 'shape/shape_properties'
module OoxmlParser
  # Class for parsing `shape`, `rect`, `oval` tags
  class Shape < OOXMLDocumentObject
    attr_accessor :type, :properties, :elements
    # @return [TextBox] text box
    attr_reader :text_box

    def initialize(type = nil,
                   properties = ShapeProperties.new,
                   elements = [],
                   parent: nil)
      @type = type
      @properties = properties
      @elements = elements
      @parent = parent
    end

    # Parse Shape object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Shape] result of parsing
    def parse(node, type)
      @type = type
      node.attribute('style').value.to_s.split(';').each do |property|
        if property.include?('margin-top')
          @properties.margins.top = property.split(':').last
        elsif property.include?('margin-left')
          @properties.margins.left = property.split(':').last
        elsif property.include?('margin-right')
          @properties.margins.right = property.split(':').last
        elsif property.include?('width')
          @properties.size.width = property.split(':').last
        elsif property.include?('height')
          @properties.size.height = property.split(':').last
        elsif property.include?('z-index')
          @properties.z_index = property.split(':').last
        elsif property.include?('position')
          @properties.position = property.split(':').last
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'textbox'
          @text_box = TextBox.new(parent: self).parse(node_child)
        end
      end
      @properties.fill_color = Color.new(parent: self).parse_hex_string(node.attribute('fillcolor').value.to_s.sub('#', '').split(' ').first) unless node.attribute('fillcolor').nil?
      @properties.stroke.weight = node.attribute('strokeweight').value unless node.attribute('strokeweight').nil?
      @properties.stroke.color = Color.new(parent: self).parse_hex_string(node.attribute('strokecolor').value.to_s.sub('#', '').split(' ').first) unless node.attribute('strokecolor').nil?
      self
    end
  end
end
