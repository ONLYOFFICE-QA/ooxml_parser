# frozen_string_literal: true

module OoxmlParser
  # Module for some helpers for ParagraphRun
  module DocxParagraphRunHelpers
    # Parse other properties
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocxParagraphRun] result of parse
    def parse_properties(node)
      self.font_style = FontStyle.new
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'rFonts'
          node_child.attributes.each do |font_attribute, value|
            case font_attribute
            when 'asciiTheme'
              theme = node_child.attribute('w:asciiTheme').value
              next unless root_object.theme

              self.font = root_object.theme.font_scheme.major_font.latin.typeface if theme.include?('major')
              self.font = root_object.theme.font_scheme.minor_font.latin.typeface if theme.include?('minor')
            when 'ascii'
              self.font = value.value
              break
            end
          end
        when 'sz'
          self.size = node_child.attribute('val').value.to_i / 2.0
        when 'highlight'
          self.highlight = node_child.attribute('val').value
        when 'vertAlign'
          self.vertical_align = node_child.attribute('val').value.to_sym
        when 'effect'
          self.effect = node_child.attribute('val').value
        when 'position'
          self.position = (node_child.attribute('val').value.to_f / (28.0 + (1.0 / 3.0)) / 2.0).round(1)
        when 'em'
          self.em = node_child.attribute('val').value
        when 'spacing'
          self.spacing = (node_child.attribute('val').value.to_f / 566.9).round(1)
        when 'textFill'
          self.text_fill = TextFill.new(parent: self).parse(node_child)
        when 'textOutline'
          self.text_outline = TextOutline.new(parent: self).parse(node_child)
        when 'bCs', 'b'
          font_style.bold = option_enabled?(node_child)
        when 'iCs', 'i'
          font_style.italic = option_enabled?(node_child)
        when 'caps'
          self.caps = :caps
        when 'smallCaps'
          self.caps = :small_caps if option_enabled?(node_child)
        when 'color'
          parse_color_tag(node_child)
        when 'shd'
          @shade = Shade.new(parent: self).parse(node_child)
        when 'u', 'uCs'
          parse_underline(node_child)
        when 'strike'
          font_style.strike = :single if option_enabled?(node_child)
        when 'dstrike'
          font_style.strike = :double if option_enabled?(node_child)
        end
      end
      self
    end

    # Temp method to return background color
    # Need to be compatible with older versions
    # @return [OoxmlParser::Color]
    def background_color
      shade.to_background_color
    end

    private

    # Parse `color` tag
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [nil]
    def parse_color_tag(node)
      color = OoxmlColor.new(parent: self).parse(node)
      node.attributes.sort.to_h.each do |key, value|
        case key
        when 'val'
          self.font_color = Color.new(parent: self).parse_hex_string(value.value)
        when 'themeColor'
          self.font_color = root_object.theme.color_scheme[value.value.to_sym].color.dup if root_object.theme && root_object.theme.color_scheme[value.value.to_sym]
        when 'themeShade'
          self.font_color = color.value if font_color.none?
          font_color.calculate_with_shade!(value.value.hex.to_f / 255.0)
        when 'themeTint'
          font_color.calculate_with_tint!(1.0 - (value.value.hex.to_f / 255.0))
        end
      end
    end

    # Parse `u` or `uCs` tag
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [nil]
    def parse_underline(node)
      font_style.underlined = Underline.new

      return unless option_enabled?(node)

      font_style.underlined.style = node.attribute('val').value
      font_style.underlined.color = Color.new(parent: font_style.underlined).parse_hex_string(node.attribute('color').value) unless node.attribute('color').nil?
    end
  end
end
