module OoxmlParser
  class ImageProperties
    attr_accessor :alpha_modulate_fixed_effect

    def self.parse(blip_properties_node)
      properties = ImageProperties.new
      blip_properties_node.xpath('*').each do |property_node|
        case property_node.name
        when 'alphaModFix'
          properties.alpha_modulate_fixed_effect = (property_node.attribute('amt').value.to_f / 1_000.0).round(1)
        end
      end
      properties
    end
  end
end
