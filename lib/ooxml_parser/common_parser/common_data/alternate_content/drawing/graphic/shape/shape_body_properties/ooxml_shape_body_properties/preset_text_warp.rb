module OoxmlParser
  class PresetTextWarp < OOXMLDocumentObject
    attr_accessor :preset

    def self.parse(node)
      preset_text_warp = PresetTextWarp.new
      node.attributes.each do |key, value|
        case key
        when 'prst'
          preset_text_warp.preset = value.value.to_sym
        end
      end
      preset_text_warp
    end
  end
end
