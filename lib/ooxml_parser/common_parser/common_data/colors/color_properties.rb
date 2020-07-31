# frozen_string_literal: true

module OoxmlParser
  # Class for color transformations
  class ColorProperties < OOXMLDocumentObject
    # @return [Integer] alpha value of color
    attr_reader :alpha
    # @return [Float] luminance modulation value
    attr_reader :luminance_modulation
    # @return [Float] luminance offset value
    attr_reader :luminance_offset
    # @return [Float] tint value
    attr_reader :tint

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
