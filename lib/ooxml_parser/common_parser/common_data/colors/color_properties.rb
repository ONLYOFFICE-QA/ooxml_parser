# frozen_string_literal: true

module OoxmlParser
  # Class for color transformations
  class ColorProperties < OOXMLDocumentObject
    # @return [ValuedChild] alpha value of color object
    attr_reader :alpha_object
    # @return [ValuedChild] luminance modulation value object
    attr_reader :luminance_modulation_object
    # @return [ValuedChild] luminance offset value object
    attr_reader :luminance_offset_object
    # @return [ValuedChild] tint value object
    attr_reader :tint_object

    # Parse ColorProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ColorProperties] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'alpha'
          @alpha_object = ValuedChild.new(:float, parent: self).parse(node_child)
        when 'lumMod'
          @luminance_modulation_object = ValuedChild.new(:float, parent: self).parse(node_child)
        when 'lumOff'
          @luminance_offset_object = ValuedChild.new(:float, parent: self).parse(node_child)
        when 'tint'
          @tint_object = ValuedChild.new(:float, parent: self).parse(node_child)
        end
      end
      self
    end

    # @return [Integer] alpha value
    def alpha
      (@alpha_object.value / 1_000.0).round
    end

    # @return [Float] luminance modulation value
    def luminance_modulation
      @luminance_modulation_object.value / 100_000.0
    end

    # @return [Float] luminance offset value
    def luminance_offset
      @luminance_offset_object.value / 100_000.0
    end

    # @return [nil, Float] tint value
    def tint
      return nil unless @tint_object

      @tint_object.value / 100_000.0
    end
  end
end
