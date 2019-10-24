# frozen_string_literal: true

require_relative 'preset_geometry/shape_adjust_value_list'
module OoxmlParser
  # Class for describing Preset Geometry
  class PresetGeometry < OOXMLDocumentObject
    # @return [String] name of preset geometry
    attr_accessor :name
    # @return [ShapeAdjustValueList] list of adjust values
    attr_accessor :adjust_values_list

    # Parse PresetGeometry
    # @param [Nokogiri::XML:Node] node with PresetGeometry
    # @return [PresetGeometry] result of parsing
    def parse(node)
      @name = node.attribute('prst').value.to_sym if node.attribute('prst')
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'avLst'
          @adjust_values_list = ShapeAdjustValueList.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
