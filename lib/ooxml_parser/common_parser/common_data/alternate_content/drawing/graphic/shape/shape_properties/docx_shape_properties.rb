# frozen_string_literal: true

require_relative 'docx_color'
require_relative 'docx_shape_size'
require_relative 'docx_shape_line'
require_relative 'docx_shape_properties/blip_fill'
require_relative 'docx_shape_properties/preset_geometry'
require_relative 'docx_shape_properties/text_box'
require_relative 'custom_geometry/ooxml_custom_geometry'
module OoxmlParser
  # DOCX Shape Properties
  class DocxShapeProperties < OOXMLDocumentObject
    # @return [DocxShapeSize] size of shape
    attr_reader :shape_size
    # @return [PresetGeometry] preset geometry of object
    attr_reader :preset_geometry
    # @return [DocxColor] color of object
    attr_reader :fill_color
    # @return [DocxShapeLine] line info
    attr_reader :line
    # @return [PresetGeometry] is some geometry custom
    attr_reader :custom_geometry
    # @return [BlipFill] BlipFill data
    attr_reader :blip_fill

    alias transform shape_size
    alias fill fill_color
    alias preset preset_geometry

    def initialize(parent: nil)
      @line = DocxShapeLine.new
      super
    end

    # Parse DocxShapeProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocxShapeProperties] result of parsing
    def parse(node)
      @fill_color = DocxColor.new(parent: self).parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'xfrm'
          @shape_size = DocxShapeSize.new(parent: self).parse(node_child)
        when 'prstGeom'
          @preset_geometry = PresetGeometry.new(parent: self).parse(node_child)
        when 'ln'
          @line = DocxShapeLine.new(parent: self).parse(node_child)
        when 'blipFill'
          @blip_fill = BlipFill.new(parent: self).parse(node_child)
        when 'custGeom'
          @preset_geometry = PresetGeometry.new(parent: self).parse(node_child)
          @preset_geometry.name = :custom
          @custom_geometry = OOXMLCustomGeometry.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
