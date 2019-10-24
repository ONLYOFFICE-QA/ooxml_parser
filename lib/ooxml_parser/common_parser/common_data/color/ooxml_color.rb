# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `color` tag
  class OoxmlColor < OOXMLDocumentObject
    # @return [Color] value of color
    attr_reader :value
    # @return [Symbol] theme color name
    attr_reader :theme_color
    # @return [Symbol] theme shade
    attr_reader :theme_shade
    # @return [Integer] theme index
    attr_reader :theme
    # @return [Float] tint
    attr_reader :tint
    # @return [Float] Indexed id
    attr_reader :indexed
    # @return [Color] rgb color
    attr_reader :rgb

    def to_color
      return Color.get_rgb_by_color_index(indexed) if indexed
      return ThemeColors.new(parent: self).parse_color_theme(theme, tint) if theme
      return rgb if rgb

      value
    end

    def ==(other)
      return to_color == other if other.is_a?(Color)
      return to_color == other if other.is_a?(Symbol)

      super
    end

    # Parse Hyperlink object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Hyperlink] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = Color.new.parse_hex_string(value.value.to_s)
        when 'themeColor'
          @theme_color = value.value.to_sym
        when 'themeShade'
          @theme_shade = Integer("0x#{value.value}")
        when 'theme'
          @theme = value.value.to_i
        when 'tint'
          @tint = value.value.to_f
        when 'indexed'
          @indexed = value.value.to_i
        when 'rgb'
          @rgb = Color.new.parse_hex_string(value.value)
        end
      end
      self
    end
  end
end
