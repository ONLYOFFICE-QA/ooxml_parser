# noinspection RubyTooManyInstanceVariablesInspection
require_relative 'docx_paragraph_run/docx_paragraph_run_helpers'
require_relative 'docx_paragraph_run/object'
require_relative 'docx_paragraph_run/text_outline'
require_relative 'docx_paragraph_run/text_fill'
require_relative 'docx_paragraph_run/shape'

module OoxmlParser
  class DocxParagraphRun < OOXMLDocumentObject
    include DocxParagraphRunHelpers
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
    # @return [RunProperties] properties of run
    attr_accessor :run_properties
    # @return [RunObject] object of run
    attr_accessor :object

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
      @object = nil
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
      character.run_properties = @run_properties
      character.object = @object
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

    def parse(r_tag, char_number, parent: nil)
      @parent = parent
      r_tag.xpath('*').each do |node_child|
        case node_child.name
        when 'rPr'
          parse_properties(node_child, DocumentStructure.default_run_style)
          @run_properties = RunProperties.new(parent: self).parse(node_child)
        when 'instrText'
          if node_child.text.include?('HYPERLINK')
            hyperlink = Hyperlink.new(node_child.text.sub('HYPERLINK ', '').split(' \\o ').first, node_child.text.sub('HYPERLINK', '').split(' \\o ').last)
            @link = hyperlink
          elsif node_child.text[/PAGE\s+\\\*/]
            @text = '*PAGE NUMBER*'
          end
        when 'fldChar'
          @fld_char = node_child.attribute('fldCharType').value.to_sym
        when 't'
          @text += node_child.text
        when 'noBreakHyphen'
          @text += '–'
        when 'tab'
          @text += "\t"
        when 'drawing'
          @drawings << DocxDrawing.parse(node_child)
        when 'AlternateContent'
          @alternate_content = AlternateContent.parse(node_child, parent: self)
        when 'br'
          if node_child.attribute('type').nil?
            @break = :line
            @text += "\r"
          else
            case node_child.attribute('type').value
            when 'page', 'column'
              @break = node_child.attribute('type').value.to_sym
            end
          end
        when 'footnoteReference'
          @footnote = HeaderFooter.parse(node_child, parent: self)
        when 'endnoteReference'
          @endnote = HeaderFooter.parse(node_child, parent: self)
        when 'pict'
          node_child.xpath('*').each do |pict_node_child|
            case pict_node_child.name
            when 'shape'
              @shape = Shape.parse(pict_node_child, :shape)
            when 'rect'
              @shape = Shape.parse(pict_node_child, :rectangle)
            when 'oval'
              @shape = Shape.parse(pict_node_child, :oval)
            end
          end
        when 'object'
          @object = RunObject.new(parent: self).parse(node_child)
        end
      end
      @number = char_number
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
  end
end
