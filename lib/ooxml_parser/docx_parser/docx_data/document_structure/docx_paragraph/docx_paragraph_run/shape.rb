# frozen_string_literal: true

require_relative 'shape/shape_properties'
module OoxmlParser
  # Class for parsing `shape`, `rect`, `oval` tags
  class Shape < OOXMLDocumentObject
    attr_accessor :type, :properties, :elements

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
      unless node.attribute('fillcolor').nil?
        @properties.fill_color = Color.new(parent: self).parse_hex_string(node.attribute('fillcolor').value.to_s.sub('#', '').split(' ').first)
      end
      @properties.stroke.weight = node.attribute('strokeweight').value unless node.attribute('strokeweight').nil?
      unless node.attribute('strokecolor').nil?
        @properties.stroke.color = Color.new(parent: self)
                                        .parse_hex_string(node.attribute('strokecolor')
                                                              .value
                                                              .to_s
                                                              .sub('#', '')
                                                              .split(' ').first)
      end
      @elements = TextBox.parse_list(node.xpath('v:textbox').first, parent: self) unless node.xpath('v:textbox').first.nil?
      self
    end
  end
end
