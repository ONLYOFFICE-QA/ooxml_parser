# frozen_string_literal: true

require_relative 'docx_custom_geometry/docx_shape_line_path'
# Docx Custom Geometry
module OoxmlParser
  class OOXMLCustomGeometry < OOXMLDocumentObject
    attr_accessor :paths_list

    def initialize(paths_list = [], parent: nil)
      @paths_list = paths_list
      @parent = parent
    end

    # Parse OOXMLCustomGeometry object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [OOXMLCustomGeometry] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'pathLst'
          node_child.xpath('a:path', 'xmlns:a' => 'http://schemas.openxmlformats.org/drawingml/2006/main').each do |path_node|
            @paths_list << DocxShapeLinePath.new(parent: self).parse(path_node)
          end
        end
      end
      self
    end
  end
end
