module OoxmlParser
  # Class for parsing `rPr` run properties
  # TODO: Merge with `RunProperties` class
  class RunPropertiesDocument < OOXMLDocumentObject
    def self.parse(character_pr_tag, character_style = DocxParagraphRun.new, default_character = DocumentStructure.default_run_style)
      character_style.font_style = FontStyle.new
      character_pr_tag.xpath('w:rStyle').each do |r_style|
        style = character_style.root_object.document_style_by_id(r_style.attribute('val').value)
        break if style.nil?
        character_style = style.run_properties
        character_style = default_character if character_style.nil?
      end
      character_pr_tag.xpath('w:rFonts').each do |r_font|
        r_font.attributes.each do |font_attribute, value|
          case font_attribute
          when 'asciiTheme'
            character_style.font = DocxParagraphRun.parse_font_by_theme(r_font.attribute('asciiTheme').value)
          when 'ascii' # , 'hAnsi', 'eastAsia', 'cs'
            character_style.font = value.value
            break
          end
        end
      end
      character_pr_tag.xpath('*').each do |run_properties_node|
        case run_properties_node.name
        when 'sz'
          character_style.size = run_properties_node.attribute('val').value.to_i / 2.0
        when 'szCs'
          character_style.font_size_complex = run_properties_node.attribute('val').value.to_i / 2.0
        when 'highlight'
          character_style.highlight = run_properties_node.attribute('val').value
        when 'vertAlign'
          character_style.vertical_align = run_properties_node.attribute('val').value.to_sym
        when 'shadow'
          character_style.shadow = true
        when 'outline'
          character_style.outline = true
        when 'imprint'
          character_style.imprint = true
        when 'emboss'
          character_style.emboss = true
        when 'vanish'
          character_style.vanish = true
        when 'effect'
          character_style.effect = run_properties_node.attribute('val').value
        when 'position'
          character_style.position = (run_properties_node.attribute('val').value.to_f / (28.0 + 1.0 / 3.0) / 2.0).round(1)
        when 'rtl'
          character_style.rtl = OOXMLDocumentObject.option_enabled?(run_properties_node)
        when 'em'
          character_style.em = run_properties_node.attribute('val').value
        when 'cs'
          character_style.cs = OOXMLDocumentObject.option_enabled?(run_properties_node)
        when 'spacing'
          character_style.spacing = (run_properties_node.attribute('val').value.to_f / 566.9).round(1)
        when 'textFill'
          character_style.text_fill = TextFill.parse(run_properties_node)
        when 'textOutline'
          character_style.text_outline = TextOutline.parse(run_properties_node)
        end
      end
      character_pr_tag.xpath('w:color').each do |color|
        color.attributes.each do |key, value|
          case key
          when 'val'
            character_style.font_color = Color.from_int16(value.value)
          when 'themeColor'
            unless ThemeColors.list[value.value.to_sym].nil?
              if value.value == 'text2' || value.value == 'background2' || value.value.include?('accent') # Don't know why. Just works
                break
              else
                character_style.font_color = ThemeColors.list[value.value.to_sym].dup
              end
            end
          when 'themeShade'
            character_style.font_color.calculate_with_shade!(value.value.hex.to_f / 255.0)
          when 'themeTint'
            character_style.font_color.calculate_with_tint!(1.0 - (value.value.hex.to_f / 255.0))
          end
        end
      end
      character_pr_tag.xpath('w:shd').each do |shd|
        character_style.background_color = shd.attribute('fill').value
        unless shd.attribute('fill').value == 'auto' || shd.attribute('fill').value == '' || shd.attribute('fill').value == 'null'
          character_style.background_color = Color.from_int16(shd.attribute('fill').value)
        end
      end
      character_pr_tag.xpath('w:uCs').each do |u|
        character_style.font_style.underlined = Underline.new
        if !u.attribute('val').nil? && u.attribute('val').value != 'none' && u.attribute('val').value != false
          character_style.font_style.underlined.style = u.attribute('val').value
          unless u.attribute('color').nil?
            character_style.font_style.underlined.color = Color.from_int16(u.attribute('color').value)
          end
        else
          character_style.font_style.underlined = Underline.new
        end
      end
      character_pr_tag.xpath('w:u').each do |u|
        character_style.font_style.underlined = Underline.new
        if !u.attribute('val').nil? && u.attribute('val').value != 'none' && u.attribute('val').value != false
          character_style.font_style.underlined.style = u.attribute('val').value
          unless u.attribute('color').nil?
            character_style.font_style.underlined.color = Color.from_int16(u.attribute('color').value)
          end
        else
          character_style.font_style.underlined = Underline.new
        end
      end
      character_pr_tag.xpath('w:bCs').each do |b|
        character_style.font_style.bold = if b.attribute('val').nil? || b.attribute('val').value != 'false'
                                            true
                                          else
                                            false
                                          end
      end
      character_pr_tag.xpath('w:b').each do |b|
        character_style.font_style.bold = if b.attribute('val').nil? || b.attribute('val').value != 'false'
                                            true
                                          else
                                            false
                                          end
      end
      character_pr_tag.xpath('w:iCs').each do |i|
        next if i.attribute('val').nil?
        character_style.font_style.italic = if i.attribute('val').value != 'false'
                                              true
                                            else
                                              false
                                            end
      end
      character_pr_tag.xpath('w:i').each do |i|
        character_style.font_style.italic = if i.attribute('val').nil? || i.attribute('val').value != 'false'
                                              true
                                            else
                                              false
                                            end
      end
      character_pr_tag.xpath('w:strike').each do |strike|
        character_style.font_style.strike = if strike.attribute('val').nil? || strike.attribute('val').value != 'false' && strike.attribute('val').value != '0'
                                              :single
                                            else
                                              :none
                                            end
      end
      character_pr_tag.xpath('w:dstrike').each do |dstrike|
        if dstrike.attribute('val').nil?
          character_style.font_style.strike = :double
        elsif dstrike.attribute('val').value == '0' || dstrike.attribute('val').value == 'false'
        else
          character_style.font_style.strike = :double
        end
      end
      character_pr_tag.xpath('w:caps').each do |caps|
        if caps.attribute('val').nil?
          character_style.caps = :caps
        elsif caps.attribute('val').value == 'false'
        else
          character_style.caps = :caps
        end
      end
      character_pr_tag.xpath('w:smallCaps').each do |small_caps|
        if small_caps.attribute('val').nil?
          character_style.caps = :small_caps
        elsif small_caps.attribute('val').value == 'false'
        else
          character_style.caps = :small_caps
        end
      end
      character_style
    end
  end
end
