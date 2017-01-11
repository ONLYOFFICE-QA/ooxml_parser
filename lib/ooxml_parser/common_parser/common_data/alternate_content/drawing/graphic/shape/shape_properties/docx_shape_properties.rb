require_relative 'docx_color'
require_relative 'docx_shape_size'
require_relative 'docx_shape_line'
require_relative 'docx_shape_properties/preset_geometry'
require_relative 'docx_shape_properties/text_box'
require_relative 'custom_geometry/ooxml_custom_geometry'
# DOCX Shape Properties
module OoxmlParser
  class DocxShapeProperties < OOXMLDocumentObject
    attr_accessor :shape_size, :preset_geometry, :fill_color, :text_box, :line, :custom_geometry

    alias transform shape_size
    alias fill fill_color
    alias preset preset_geometry

    def initialize(parent: nil)
      @line = DocxShapeLine.new
      @parent = parent
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
        when 'txbx'
          @text_box = TextBox.parse_list(node_child)
        when 'ln'
          @line = DocxShapeLine.new(parent: self).parse(node_child)
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
