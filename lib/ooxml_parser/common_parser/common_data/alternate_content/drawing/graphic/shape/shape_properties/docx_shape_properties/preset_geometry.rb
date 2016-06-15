require_relative 'preset_geometry/shape_adjust_value_list'
module OoxmlParser
  # Class for describing Preset Geometry
  class PresetGeometry
    # @return [String] name of preset geometry
    attr_accessor :name
    # @return [ShapeAdjustValueList] list of adjust values
    attr_accessor :adjust_values_list

    # Parse Relationships
    # @param [Nokogiri::XML:Node] node with PresetGeometry
    # @return [PresetGeometry] result of parsing
    def self.parse(node)
      geometry = PresetGeometry.new
      geometry.name = node.attribute('prst').value.to_sym if node.attribute('prst')
      node.xpath('*').each do |preset_geometry_child|
        case preset_geometry_child.name
        when 'avLst'
          geometry.adjust_values_list = ShapeAdjustValueList.parse(preset_geometry_child)
        end
      end
      geometry
    end
  end
end
