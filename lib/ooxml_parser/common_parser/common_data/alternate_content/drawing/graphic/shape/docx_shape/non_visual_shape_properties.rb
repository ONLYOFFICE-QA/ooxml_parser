require_relative 'non_visual_shape_properties/common_non_visual_properties'
require_relative 'non_visual_shape_properties/non_visual_properties'
module OoxmlParser
  class NonVisualShapeProperties < OOXMLDocumentObject
    attr_accessor :common_properties, :non_visual_properties

    # Parse NonVisualShapeProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [NonVisualShapeProperties] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'cNvPr'
          @common_properties = CommonNonVisualProperties.new(parent: self).parse(node_child)
        when 'nvPr'
          @non_visual_properties = NonVisualProperties.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
