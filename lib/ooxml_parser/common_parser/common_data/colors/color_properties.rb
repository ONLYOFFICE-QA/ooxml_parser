module OoxmlParser
  # Class for color transformations
  class ColorProperties < OOXMLDocumentObject
    attr_accessor :alpha, :luminance_modulation, :luminance_offset
    attr_accessor :tint

    # Parse ColorProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ColorProperties] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'alpha'
          @alpha = (node_child.attribute('val').value.to_f / 1_000.0).round
        when 'lumMod'
          @luminance_modulation = node_child.attribute('val').value.to_f / 100_000.0
        when 'lumOff'
          @luminance_offset = node_child.attribute('val').value.to_f / 100_000.0
        when 'tint'
          @tint = node_child.attribute('val').value.to_f / 100_000.0
        end
      end
      self
    end
  end
end
