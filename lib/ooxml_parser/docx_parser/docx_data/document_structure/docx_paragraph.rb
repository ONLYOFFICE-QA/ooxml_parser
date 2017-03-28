# noinspection RubyTooManyInstanceVariablesInspection
require_relative 'docx_paragraph/bookmark'
require_relative 'docx_paragraph/docx_paragraph_helper'
require_relative 'docx_paragraph/docx_paragraph_run'
require_relative 'docx_paragraph/indents'
require_relative 'docx_paragraph/inserted'
require_relative 'docx_paragraph/frame_properties'
require_relative 'docx_paragraph/docx_formula'
require_relative 'docx_paragraph/style_parametres'
module OoxmlParser
  class DocxParagraph < OOXMLDocumentObject
    include DocxParagraphHelper
    attr_accessor :number, :bookmark_start, :bookmark_end, :align, :spacing, :background_color, :ind, :numbering,
                  :character_style_array, :horizontal_line, :page_break, :kinoku, :borders, :keep_lines,
                  :contextual_spacing, :sector_properties, :page_numbering, :section_break, :style, :keep_next,
                  :orphan_control, :frame_properties
    # @return [ParagraphProperties] Properties of current paragraph
    attr_accessor :paragraph_properties
    # @return [Inserted] data inserted by review
    attr_accessor :inserted
    # @return [Integer] id of paragraph (for comment)
    attr_accessor :paragraph_id
    # @return [Integer] id of text (for comment)
    attr_accessor :text_id

    def initialize(parent: nil)
      @number = 0
      @bookmark_start = []
      @bookmark_end = []
      @align = 'left'
      @spacing = Spacing.new
      @background_color = nil
      @ind = Indents.new
      @kinoku = false
      @numbering = nil
      @character_style_array = []
      @horizontal_line = false
      @page_break = false
      @borders = Borders.new
      @keep_lines = false
      @contextual_spacing = false
      @sector_properties = nil
      @page_numbering = false
      @section_break = nil
      @style = nil
      @keep_next = false
      @orphan_control = true
      @frame_properties = nil
      @parent = parent
    end

    def copy
      paragraph = DocxParagraph.new
      paragraph.number = number
      paragraph.bookmark_start = @bookmark_start.dup
      paragraph.bookmark_end = @bookmark_end.dup
      paragraph.align = @align
      paragraph.spacing = @spacing.copy
      paragraph.background_color = @background_color
      paragraph.ind = @ind.dup
      paragraph.numbering = @numbering
      paragraph.character_style_array = @character_style_array
      paragraph.horizontal_line = @horizontal_line
      paragraph.page_break = @page_break
      paragraph.kinoku = @kinoku
      paragraph.borders = @borders
      paragraph.keep_lines = @keep_lines
      paragraph.contextual_spacing = @contextual_spacing
      paragraph.sector_properties = @sector_properties
      paragraph.page_numbering = @page_numbering
      paragraph.section_break = @section_break
      paragraph.style = @style
      paragraph.keep_next = @keep_next
      paragraph.orphan_control = @orphan_control
      paragraph.frame_properties = @frame_properties
      paragraph.paragraph_properties = @paragraph_properties
      paragraph.inserted = @inserted
      paragraph.paragraph_id = @paragraph_id
      paragraph.text_id = @text_id
      paragraph.parent = @parent
      paragraph
    end

    def nonempty_runs
      @character_style_array.select do |cur_run|
        if cur_run.is_a?(DocxParagraphRun)
          (!cur_run.text.empty? ||
              !cur_run.alternate_content.nil? ||
              !cur_run.drawing.nil? ||
              !cur_run.object.nil? ||
              !cur_run.footnote.nil? ||
              !cur_run.endnote.nil?
          )
        elsif cur_run.is_a?(DocxFormula)
          true
        end
      end
    end

    # @return [True, false] if structure contain any user data
    def with_data?
      !nonempty_runs.empty?
    end

    def remove_empty_runs
      nonempty = nonempty_runs
      @character_style_array.each do |cur_run|
        @character_style_array.delete(cur_run) unless nonempty.include?(cur_run)
      end
    end

    def ==(other)
      ignored_attributes = [:@number, :@parent]
      all_instance_variables = instance_variables
      significan_attribues = all_instance_variables - ignored_attributes
      significan_attribues.each do |current_attributes|
        unless instance_variable_get(current_attributes) == other.instance_variable_get(current_attributes)
          return false
        end
      end
      true
    end

    def parse(node, par_number = 0, default_character = DocxParagraphRun.new, parent: nil)
      @parent = parent
      default_character_style = default_character.copy
      character_styles_array = []
      custom_character_style = default_character_style.copy
      char_number = 0
      comments = []
      node.attributes.each do |key, value|
        case key
        when 'paraId'
          @paragraph_id = value.value.to_i
        when 'textId'
          @text_id = value.value.to_i
        end
      end
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'bookmarkStart'
          @bookmark_start << Bookmark.new(parent: self).parse(node_child)
        when 'bookmarkEnd'
          @bookmark_end << Bookmark.new(parent: self).parse(node_child)
        when 'pPr'
          parse_paragraph_style(node_child, custom_character_style)
          node.xpath('w:pict').each do |pict|
            pict.xpath('v:rect').each do
              @horizontal_line = true
            end
          end
          @paragraph_properties = ParagraphProperties.new(parent: self).parse(node_child)
        when 'commentRangeStart'
          comments << node_child.attribute('id').value
        when 'fldSimple'
          instruction = node_child.attribute('instr').to_s
          @page_numbering = true if instruction.include?('PAGE')
          node_child.xpath('w:r').each do |r_tag|
            character_style = default_character_style.copy
            character_style.parse(r_tag, char_number, parent: parent)
            character_style.page_number = @page_numbering
            character_style.instruction = instruction
            character_styles_array << character_style.copy
            char_number += 1
          end
        when 'r'
          character_style = custom_character_style.copy
          node_child.xpath('w:instrText').each do |insrt_text|
            @page_numbering = true if insrt_text.text.include?('PAGE')
          end
          character_style.parse(node_child, char_number, parent: self)
          character_style.comments = comments.dup
          character_styles_array << character_style.copy
          unless character_style.shape.nil?
            character_styles_array.last.shape = character_style.shape
          end
          char_number += 1
        when 'hyperlink'
          character_style = default_character_style.copy
          if !node_child.attribute('id').nil?
            character_style.link = Hyperlink.new(parent: character_style).parse(node_child)
          else
            unless node_child.attribute('anchor').nil?
              character_style.link = node_child.attribute('anchor').value
            end
          end
          node_child.xpath('w:r').each do |r_tag|
            character_style.parse(r_tag, char_number, parent: parent)
            character_styles_array << character_style.copy
            char_number += 1
          end
          node_child.xpath('w:fldSimple').each do |simple_field|
            instruction = simple_field.attribute('instr').to_s
            @page_numbering = true if instruction.include?('PAGE')
            simple_field.xpath('w:r').each do |r_tag|
              character_style.parse(r_tag, char_number, parent: self)
              character_style.page_number = @page_numbering
              character_style.instruction = instruction
              character_styles_array << character_style.copy
              char_number += 1
            end
          end
        when 'oMathPara'
          node_child.xpath('m:oMath').each do |o_math|
            character_styles_array << DocxFormula.new(parent: self).parse(o_math)
          end
        when 'commentRangeEnd'
          comments.each_with_index do |comment, index|
            if comment == node_child.attribute('id').value
              comments.delete_at(index)
              break
            end
          end
        when 'ins'
          @inserted = Inserted.new(parent: self).parse(node_child)
        end
      end
      @number = par_number
      if character_styles_array.last.class == DocxParagraphRun
        character_styles_array.last.text = character_styles_array.last.text.rstrip
      end
      @character_style_array = character_styles_array
      @parent = parent
      self
    end

    def parse_paragraph_style(node, default_char_style = DocxParagraphRun.new)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'pageBreakBefore'
          if node_child.attribute('val').nil? || node_child.attribute('val').value != 'false'
            @page_break = true
          end
        when 'pBdr'
          @borders = ParagraphBorders.new(parent: self).parse(node_child)
        when 'keepLines'
          if node_child.attribute('val').nil?
            @keep_lines = true
          else
            unless node_child.attribute('val').value == 'false'
              @keep_lines = true
            end
          end
        when 'widowControl'
          @orphan_control = option_enabled?(node_child)
        when 'keepNext'
          @keep_next = true
        when 'contextualSpacing'
          @contextual_spacing = true
        when 'shd'
          background_color_string = node_child.attribute('fill').value
          @background_color = Color.new(parent: self).parse_hex_string(background_color_string)
          unless node_child.attribute('val').nil?
            @background_color.set_style(node_child.attribute('val').value)
          end
        when 'pStyle'
          DocxParagraph.parse_paragraph_style_xml(node_child.attribute('val').value, self, default_char_style)
        when 'ind'
          @ind = DocumentStructure.default_paragraph_style.ind.dup.parse(node_child)
        when 'kinoku'
          @kinoku = true
        when 'framePr'
          @frame_properties = FrameProperties.new(parent: self).parse(node_child)
        when 'numPr'
          @numbering = NumberingProperties.new(parent: self).parse(node_child)
        when 'jc'
          @align = node_child.attribute('val').value.to_sym unless node_child.attribute('val').nil?
          @align = :justify if node_child.attribute('val').value == 'both'
        when 'spacing'
          unless node_child.attribute('before').nil?
            @spacing.before = (node_child.attribute('before').value.to_f / 566.9).round(2)
          end
          unless node_child.attribute('after').nil?
            @spacing.after = (node_child.attribute('after').value.to_f / 566.9).round(2)
          end
          unless node_child.attribute('lineRule').nil?
            @spacing.line_rule = node_child.attribute('lineRule').value.sub('atLeast', 'at_least').to_sym
          end
          unless node_child.attribute('line').nil?
            @spacing.line = (@spacing.line_rule == :auto ? (node_child.attribute('line').value.to_f / 240.0).round(2) : (node_child.attribute('line').value.to_f / 566.9).round(2))
          end
        when 'sectPr'
          @sector_properties = PageProperties.new(parent: self).parse(node_child, self, default_char_style)
          @section_break = case @sector_properties.type
                           when 'oddPage'
                             'Odd page'
                           when 'evenPage'
                             'Even page'
                           when 'continuous'
                             'Current Page'
                           else
                             'Next Page'
                           end
        end
      end
      @parent = parent
      self
    end

    def self.parse_paragraph_style_xml(id, paragraph_style, character_style)
      doc = Nokogiri::XML(File.open(OOXMLDocumentObject.path_to_folder + 'word/styles.xml'))
      doc.search('//w:style').each do |style|
        next unless style.attribute('styleId').value == id
        style.xpath('w:pPr').each do |p_pr|
          paragraph_style.parse_paragraph_style(p_pr, character_style)
          paragraph_style.style = StyleParametres.new(parent: paragraph_style).parse(style)
        end
        style.xpath('w:rPr').each do |r_pr|
          character_style.parse_properties(r_pr, DocumentStructure.default_run_style)
        end
        break
      end
    end
  end
end
