# frozen_string_literal: true

require_relative 'docx_drawing/doc_properties'
require_relative 'drawing_properties/docx_drawing_properties'
require_relative 'drawing_properties/size_relative_horizontal'
require_relative 'drawing_properties/size_relative_vertical'
require_relative 'graphic/docx_graphic'
module OoxmlParser
  # Class for parsing `graphic` tags
  class DocxDrawing < OOXMLDocumentObject
    attr_accessor :type, :properties, :graphic
    # @return [DocProperties] doc properties
    attr_accessor :doc_properties

    alias picture graphic

    def initialize(properties = DocxDrawingProperties.new, parent: nil)
      @properties = properties
      @parent = parent
    end

    # Parse DocxDrawing
    # @param [Nokogiri::XML:Node] node with NumberingProperties
    # @return [DocxDrawing] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'anchor'
          @type = :flow
        when 'inline'
          @type = :inline
        end
        @properties.distance_from_text = DocxDrawingDistanceFromText.new(parent: self).parse(node_child)
        @properties.wrap = DocxWrapDrawing.new(parent: self).parse(node_child)
        node_child.attributes.each do |key, value|
          case key
          when 'relativeHeight'
            @properties.relative_height = value.value.to_i
          end
        end
        node_child.xpath('*').each do |content_node_child|
          case content_node_child.name
          when 'simplePos'
            @properties.simple_position = OOXMLCoordinates.parse(content_node_child)
          when 'extent'
            @properties.object_size = OOXMLCoordinates.parse(content_node_child, x_attr: 'cx', y_attr: 'cy', unit: :emu)
          when 'graphic'
            @graphic = DocxGraphic.new(parent: self).parse(content_node_child)
          when 'positionV'
            @properties.vertical_position = DocxDrawingPosition.new(parent: self).parse(content_node_child)
          when 'positionH'
            @properties.horizontal_position = DocxDrawingPosition.new(parent: self).parse(content_node_child)
          when 'sizeRelH'
            @properties.size_relative_horizontal = SizeRelativeHorizontal.new(parent: self).parse(content_node_child)
          when 'sizeRelV'
            @properties.size_relative_vertical = SizeRelativeVertical.new(parent: self).parse(content_node_child)
          when 'docPr'
            @doc_properties = DocProperties.new(parent: self).parse(content_node_child)
          end
        end
      end
      self
    end
  end
end
