# frozen_string_literal: true

require_relative 'ooxml_coordinates'
require_relative 'ooxml_size'
require_relative 'docx_drawing_distance_from_text'
require_relative 'docx_drawing_position'
require_relative 'docx_wrap_drawing'
module OoxmlParser
  # Docx Drawing Properties
  class DocxDrawingProperties < OOXMLDocumentObject
    # @return [DocxDrawingDistanceFromText] distance from text
    attr_reader :distance_from_text
    # @return [DocxWrapDrawing] wrap of drawing
    attr_reader :wrap
    # @return [Integer] relative height of object
    attr_reader :relative_height
    # @return [OOXMLCoordinates] simple position of object
    attr_reader :simple_position
    # @return [OOXMLCoordinates] size of object
    attr_reader :object_size
    # @return [DocxDrawingPosition] vertical position
    attr_reader :vertical_position
    # @return [DocxDrawingPosition] horizontal position
    attr_reader :horizontal_position
    # @return [SizeRelativeHorizontal] size of drawing relative to horizontal
    attr_reader :size_relative_horizontal
    # @return [SizeRelativeVertical] size of drawing relative to vertical
    attr_reader :size_relative_vertical

    # Parse DocxDrawingProperties
    # @param [Nokogiri::XML:Node] node with DocxDrawingProperties
    # @return [DocxDrawingProperties] result of parsing
    def parse(node)
      @distance_from_text = DocxDrawingDistanceFromText.new(parent: self).parse(node)
      @wrap = DocxWrapDrawing.new(parent: self).parse(node)

      node.attributes.each do |key, value|
        case key
        when 'relativeHeight'
          @relative_height = value.value.to_i
        end
      end

      node.xpath('*').each do |content_node_child|
        case content_node_child.name
        when 'simplePos'
          @simple_position = OOXMLCoordinates.new(parent: self).parse(content_node_child)
        when 'extent'
          @object_size = OOXMLCoordinates.new(parent: self)
                                         .parse(content_node_child,
                                                x_attr: 'cx',
                                                y_attr: 'cy',
                                                unit: :emu)
        when 'positionV'
          @vertical_position = DocxDrawingPosition.new(parent: self).parse(content_node_child)
        when 'positionH'
          @horizontal_position = DocxDrawingPosition.new(parent: self).parse(content_node_child)
        when 'sizeRelH'
          @size_relative_horizontal = SizeRelativeHorizontal.new(parent: self).parse(content_node_child)
        when 'sizeRelV'
          @size_relative_vertical = SizeRelativeVertical.new(parent: self).parse(content_node_child)
        end
      end
      self
    end
  end
end
