require_relative 'docx_custom_geometry/docx_shape_line_path'
# Docx Custom Geometry
module OoxmlParser
  class OOXMLCustomGeometry
    attr_accessor :paths_list

    def initialize(paths_list = [])
      @paths_list = paths_list
    end

    def self.parse(custom_geometry_node)
      custom_geometry = OOXMLCustomGeometry.new
      custom_geometry_node.xpath('*').each do |custom_geometry_node_child|
        case custom_geometry_node_child.name
        when 'pathLst'
          custom_geometry_node_child.xpath('a:path', 'xmlns:a' => 'http://schemas.openxmlformats.org/drawingml/2006/main').each { |path_node| custom_geometry.paths_list << DocxShapeLinePath.parse(path_node) }
        end
      end
      custom_geometry
    end
  end
end
