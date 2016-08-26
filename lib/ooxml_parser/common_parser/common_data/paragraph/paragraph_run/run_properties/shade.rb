module OoxmlParser
  # Class for parsing `w:shd` object
  class Shade < OOXMLDocumentObject
    # @return [Symbol] value of shade
    attr_accessor :value
    # @return [Symbol] color of shade
    attr_accessor :color
    # @return [Color] fill of shade
    attr_accessor :fill

    # Parse Shade
    # @param [Nokogiri::XML:Node] node with Shade
    # @return [Shade] result of parsing
    def self.parse(node)
      shade = Shade.new
      node.attributes.each do |key, value|
        case key
        when 'val'
          shade.value = value.value.to_sym
        when 'color'
          shade.color = value.value.to_sym
        when 'fill'
          shade.fill = Color.from_int16(value.value.to_s)
        end
      end
      shade
    end
  end
end
