module OoxmlParser
  # Class for parsing `prstTxWarp` tags
  class PresetTextWarp < OOXMLDocumentObject
    attr_accessor :preset

    # Parse PresetTextWarp
    # @param [Nokogiri::XML:Node] node with PresetTextWarp
    # @return [PresetTextWarp] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'prst'
          @preset = value.value.to_sym
        end
      end
      self
    end
  end
end
