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
      super(parent: parent)
    end

    # Parse Shape object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Shape] result of parsing
    def parse(node, type)
      @type = type
      node.attributes.each do |key, value|
        case key
        when 'fillcolor'
          @properties.fill_color = Color.new(parent: self)
                                        .parse_hex_string(value_to_hex(value))
        when 'strokecolor'
          @properties.stroke.color = Color.new(parent: self)
                                          .parse_hex_string(value_to_hex(value))
        when 'strokeweight'
          @properties.stroke.weight = value.value
        when 'style'
          parse_style(value.value)
        end
      end
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'textbox'
          @elements = TextBox.parse_list(node_child, parent: self)
        end
      end
      self
    end

    private

    # @param style [String] value to parse
    def parse_style(style)
      style.split(';').each do |property|
        if property.include?('margin-top')
          @properties.margins.top = property.split(':').last
        elsif property.include?('margin-left')
          @properties.margins.left = property.split(':').last
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
    end

    # @param value [String] value to extract hex string
    # @return [String] hex string value
    def value_to_hex(value)
      value.to_s
           .sub('#', '')
           .split
           .first
    end
  end
end
