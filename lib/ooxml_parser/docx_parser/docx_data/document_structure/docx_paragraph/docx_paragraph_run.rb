# noinspection RubyTooManyInstanceVariablesInspection
require_relative 'docx_paragraph_run/text_outline'
require_relative 'docx_paragraph_run/text_fill'
require_relative 'docx_paragraph_run/shape'

module OoxmlParser
  class DocxParagraphRun < OOXMLDocumentObject
    attr_accessor :number, :font, :vertical_align, :size, :font_color, :background_color, :font_style, :text, :drawings,
                  :link, :highlight, :shadow, :outline, :imprint, :emboss, :vanish, :effect, :caps, :w,
                  :position, :rtl, :em, :cs, :spacing, :break, :touch, :shape, :footnote, :endnote, :fld_char, :style,
                  :comments, :alternate_content, :page_number, :text_outline, :text_fill
    # @return [Float]
    # This element specifies the font size which shall be applied to all
    # complex script characters in the contents of this run when displayed
    attr_accessor :font_size_complex

    # @return [String] type of instruction used for upper level of run
    # http://officeopenxml.com/WPfieldInstructions.php
    attr_accessor :instruction

    def initialize
      @number = 0
      @font = 'Arial'
      @vertical_align = :baseline
      @size = 11
      @font_color = Color.new
      @background_color = nil
      @font_style = FontStyle.new
      @text = ''
      @drawings = []
      @link = nil
      @highlight = nil
      @shadow = nil
      @outline = nil
      @imprint = nil
      @emboss = nil
      @vanish = nil
      @effect = nil
      @caps = nil
      @w = false
      @position = 0.0
      @rtl = false
      @em = nil
      @cs = false
      @spacing = 0.0
      @break = false
      @touch = nil
      @footnote = nil
      @endnote = nil
      @fld_char = nil
      @style = nil
      @comments = []
      @alternate_content = nil
      @page_number = false
      @text_outline = nil
      @text_fill = nil
      @instruction = nil
    end

    def drawing
      # TODO: Rewrite all tests without this methos
      @drawings.empty? ? nil : drawings.first
    end

    def copy
      character = DocxParagraphRun.new
      character.number = number
      character.font = font
      character.vertical_align = vertical_align
      character.size = size
      character.font_color = font_color
      character.background_color = @background_color
      character.font_style = @font_style
      character.text = @text
      character.drawings = @drawings.clone
      character.link = @link
      character.highlight = @highlight
      character.shadow = @shadow
      character.outline = @outline
      character.imprint = @imprint
      character.emboss = @emboss
      character.vanish = @vanish
      character.effect = @effect
      character.caps = @caps
      character.w = @w
      character.position = @position
      character.rtl = @rtl
      character.em = @em
      character.cs = @cs
      character.spacing = @spacing
      character.cs = @cs
      character.break = @break
      character.touch = @touch
      character.footnote = @footnote
      character.endnote = @endnote
      character.fld_char = @fld_char
      character.style = @style
      character.comments = @comments.clone
      character.alternate_content = @alternate_content
      character.page_number = @page_number
      character.text_outline = @text_outline
      character.text_fill = @text_fill
      character.instruction = @instruction
      character
    end

    def ==(other)
      ignored_attributes = [:@number]
      all_instance_variables = instance_variables
      significan_attribues = all_instance_variables - ignored_attributes
      significan_attribues.each do |current_attributes|
        unless instance_variable_get(current_attributes) == other.instance_variable_get(current_attributes)
          return false
        end
      end
      true
    end

    def self.parse(character_pr_tag, character_style = DocxParagraphRun.new, default_character = DocumentStructure.default_run_style)
      character_style.font_style = FontStyle.new
      character_pr_tag.xpath('w:rStyle').each do |r_style|
        style = get_style_by_id(r_style.attribute('val').value)
        break if style.nil?
        style.xpath('w:rPr').each do |style_rpr|
          character_style = parse(style_rpr, character_style, default_character)
        end
      end
      character_pr_tag.xpath('w:rFonts').each do |r_font|
        r_font.attributes.each do |font_attribute, value|
          case font_attribute
          when 'asciiTheme'
            character_style.font = parse_font_by_theme(r_font.attribute('asciiTheme').value)
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

    def self.parse_character(r_tag, character_style, char_number)
      r_tag.xpath('*').each do |r_node_child|
        case r_node_child.name
        when 'rPr'
          DocxParagraphRun.parse(r_node_child, character_style, DocumentStructure.default_run_style)
        when 'instrText'
          if r_node_child.text.include?('HYPERLINK')
            hyperlink = Hyperlink.new(r_node_child.text.sub('HYPERLINK ', '').split(' \\o ').first, r_node_child.text.sub('HYPERLINK', '').split(' \\o ').last)
            character_style.link = hyperlink
          elsif r_node_child.text[/PAGE\s+\\\*/]
            character_style.text = '*PAGE NUMBER*'
          end
        when 'fldChar'
          character_style.fld_char = r_node_child.attribute('fldCharType').value.to_sym
        when 't'
          character_style.text += r_node_child.text
        when 'noBreakHyphen'
          character_style.text += '–'
        when 'tab'
          character_style.text += "\t"
        when 'drawing'
          character_style.drawings << DocxDrawing.parse(r_node_child)
        when 'AlternateContent'
          character_style.alternate_content = AlternateContent.parse(r_node_child)
        when 'br'
          if r_node_child.attribute('type').nil?
            character_style.break = :line
            character_style.text += "\r"
          else
            case r_node_child.attribute('type').value
            when 'page', 'column'
              character_style.break = r_node_child.attribute('type').value.to_sym
            end
          end
        when 'footnoteReference'
          character_style.footnote = HeaderFooter.parse(r_node_child)
        when 'endnoteReference'
          character_style.endnote = HeaderFooter.parse(r_node_child)
        when 'pict'
          r_node_child.xpath('*').each do |pict_node_child|
            case pict_node_child.name
            when 'shape'
              character_style.shape = Shape.parse(pict_node_child, :shape)
            when 'rect'
              character_style.shape = Shape.parse(pict_node_child, :rectangle)
            when 'oval'
              character_style.shape = Shape.parse(pict_node_child, :oval)
            when 'shapetype'
            end
          end
        end
      end
      character_style.number = char_number
      character_style
    end

    def self.parse_font_by_theme(theme)
      doc = Nokogiri::XML(File.open("#{OOXMLDocumentObject.path_to_folder}/#{OOXMLDocumentObject.root_subfolder}/theme/theme1.xml"))
      doc.search('//a:fontScheme').each do |font_scheme|
        if theme.include?('major')
          font_scheme.xpath('a:majorFont').each do |major_font|
            major_font.xpath('a:latin').each do |latin|
              return latin.attribute('typeface').value
            end
          end
        elsif theme.include?('minor')
          font_scheme.xpath('a:minorFont').each do |minor_font|
            minor_font.xpath('a:latin').each do |latin|
              return latin.attribute('typeface').value
            end
          end
        end
      end
    end

    def self.get_style_by_id(id)
      doc = Nokogiri::XML(File.open("#{OOXMLDocumentObject.path_to_folder}/#{OOXMLDocumentObject.root_subfolder}/styles.xml"))
      doc.search('//w:styles').each do |styles|
        styles.xpath('w:style').each do |style|
          return style if style.attribute('styleId').value == id
        end
      end
      nil
    end
  end
end
