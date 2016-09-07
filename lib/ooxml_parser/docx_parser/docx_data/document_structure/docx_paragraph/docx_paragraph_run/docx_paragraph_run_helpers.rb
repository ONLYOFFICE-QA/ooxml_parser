module OoxmlParser
  module DocxParagraphRunHelpers
    def parse_properties(character_pr_tag, _default_character = DocumentStructure.default_run_style)
      self.font_style = FontStyle.new
      character_pr_tag.xpath('w:rFonts').each do |r_font|
        r_font.attributes.each do |font_attribute, value|
          case font_attribute
          when 'asciiTheme'
            self.font = DocxParagraphRun.parse_font_by_theme(r_font.attribute('asciiTheme').value)
          when 'ascii' # , 'hAnsi', 'eastAsia', 'cs'
            self.font = value.value
            break
          end
        end
      end
      character_pr_tag.xpath('*').each do |run_properties_node|
        case run_properties_node.name
        when 'sz'
          self.size = run_properties_node.attribute('val').value.to_i / 2.0
        when 'szCs'
          self.font_size_complex = run_properties_node.attribute('val').value.to_i / 2.0
        when 'highlight'
          self.highlight = run_properties_node.attribute('val').value
        when 'vertAlign'
          self.vertical_align = run_properties_node.attribute('val').value.to_sym
        when 'shadow'
          self.shadow = true
        when 'outline'
          self.outline = true
        when 'imprint'
          self.imprint = true
        when 'emboss'
          self.emboss = true
        when 'vanish'
          self.vanish = true
        when 'effect'
          self.effect = run_properties_node.attribute('val').value
        when 'position'
          self.position = (run_properties_node.attribute('val').value.to_f / (28.0 + 1.0 / 3.0) / 2.0).round(1)
        when 'rtl'
          self.rtl = option_enabled?(run_properties_node)
        when 'em'
          self.em = run_properties_node.attribute('val').value
        when 'cs'
          self.cs = option_enabled?(run_properties_node)
        when 'spacing'
          self.spacing = (run_properties_node.attribute('val').value.to_f / 566.9).round(1)
        when 'textFill'
          self.text_fill = TextFill.parse(run_properties_node)
        when 'textOutline'
          self.text_outline = TextOutline.new(parent: self).parse(run_properties_node)
        when 'bCs', 'b'
          font_style.bold = option_enabled?(run_properties_node)
        when 'iCs', 'i'
          font_style.italic = option_enabled?(run_properties_node)
        when 'caps'
          self.caps = :caps
        when 'smallCaps'
          self.caps = :small_caps if option_enabled?(run_properties_node)
        end
      end
      character_pr_tag.xpath('w:color').each do |color|
        color.attributes.each do |key, value|
          case key
          when 'val'
            self.font_color = Color.from_int16(value.value)
          when 'themeColor'
            unless ThemeColors.list[value.value.to_sym].nil?
              if value.value == 'text2' || value.value == 'background2' || value.value.include?('accent') # Don't know why. Just works
                break
              else
                self.font_color = ThemeColors.list[value.value.to_sym].dup
              end
            end
          when 'themeShade'
            font_color.calculate_with_shade!(value.value.hex.to_f / 255.0)
          when 'themeTint'
            font_color.calculate_with_tint!(1.0 - (value.value.hex.to_f / 255.0))
          end
        end
      end
      character_pr_tag.xpath('w:shd').each do |shd|
        self.background_color = shd.attribute('fill').value
        unless shd.attribute('fill').value == 'auto' || shd.attribute('fill').value == '' || shd.attribute('fill').value == 'null'
          self.background_color = Color.from_int16(shd.attribute('fill').value)
        end
      end
      character_pr_tag.xpath('w:uCs').each do |u|
        font_style.underlined = Underline.new
        if !u.attribute('val').nil? && u.attribute('val').value != 'none' && u.attribute('val').value != false
          font_style.underlined.style = u.attribute('val').value
          unless u.attribute('color').nil?
            font_style.underlined.color = Color.from_int16(u.attribute('color').value)
          end
        else
          font_style.underlined = Underline.new
        end
      end
      character_pr_tag.xpath('w:u').each do |u|
        font_style.underlined = Underline.new
        if !u.attribute('val').nil? && u.attribute('val').value != 'none' && u.attribute('val').value != false
          font_style.underlined.style = u.attribute('val').value
          unless u.attribute('color').nil?
            font_style.underlined.color = Color.from_int16(u.attribute('color').value)
          end
        else
          font_style.underlined = Underline.new
        end
      end
      character_pr_tag.xpath('w:strike').each do |strike|
        font_style.strike = if strike.attribute('val').nil? || strike.attribute('val').value != 'false' && strike.attribute('val').value != '0'
                              :single
                            else
                              :none
                            end
      end
      character_pr_tag.xpath('w:dstrike').each do |dstrike|
        if dstrike.attribute('val').nil?
          font_style.strike = :double
        elsif dstrike.attribute('val').value == '0' || dstrike.attribute('val').value == 'false'
        else
          font_style.strike = :double
        end
      end
      self
    end
  end
end
