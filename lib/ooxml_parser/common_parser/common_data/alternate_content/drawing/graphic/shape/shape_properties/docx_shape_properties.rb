require_relative 'docx_color'
require_relative 'docx_shape_size'
require_relative 'docx_shape_line'
require_relative 'docx_shape_properties/text_box'
require_relative 'custom_geometry/ooxml_custom_geometry'
# DOCX Shape Properties
module OoxmlParser
  class DocxShapeProperties < OOXMLDocumentObject
    attr_accessor :shape_size, :preset_geometry, :fill_color, :text_box, :line, :custom_geometry

    alias_method :transform, :shape_size
    alias_method :fill, :fill_color
    alias_method :preset, :preset_geometry

    def initialize
      @line = DocxShapeLine.new
    end

    def self.parse(shape_properties_node)
      shape_properties = DocxShapeProperties.new
      shape_properties.fill_color = DocxColor.parse(shape_properties_node)
      shape_properties_node.xpath('*').each do |shape_properties_node_child|
        case shape_properties_node_child.name
        when 'xfrm'
          shape_properties.shape_size = DocxShapeSize.parse(shape_properties_node_child)
        when 'prstGeom'
          shape_properties.preset_geometry = shape_properties_node_child.attribute('prst').value.to_sym
        when 'txbx'
          shape_properties.text_box = TextBox.parse_list(shape_properties_node_child)
        when 'ln'
          shape_properties.line = DocxShapeLine.parse(shape_properties_node_child)
        when 'custGeom'
          shape_properties.preset_geometry = :custom
          shape_properties.custom_geometry = OOXMLCustomGeometry.parse(shape_properties_node_child)
        end
      end
      shape_properties
    end
  end
end
