# noinspection RubyTooManyInstanceVariablesInspection
require_relative 'docx_paragraph/bookmark'
require_relative 'docx_paragraph/docx_paragraph_helper'
require_relative 'docx_paragraph/docx_paragraph_run'
require_relative 'docx_paragraph/field_simple'
require_relative 'docx_paragraph/indents'
require_relative 'docx_paragraph/inserted'
require_relative 'docx_paragraph/structured_document_tag'
require_relative 'docx_paragraph/frame_properties'
require_relative 'docx_paragraph/docx_formula'
require_relative 'docx_paragraph/style_parametres'
module OoxmlParser
  class DocxParagraph < OOXMLDocumentObject
    include DocxParagraphHelper
    attr_accessor :number, :bookmark_start, :bookmark_end, :align, :spacing, :background_color, :ind, :numbering,
                  :character_style_array, :horizontal_line, :page_break, :borders, :keep_lines,
                  :contextual_spacing, :sector_properties, :page_numbering, :section_break, :style, :keep_next,
                  :orphan_control
    # @return [Hyperlink] hyperlink in paragraph
    attr_reader :hyperlink
    # @return [ParagraphProperties] Properties of current paragraph
    attr_accessor :paragraph_properties
    # @return [Inserted] data inserted by review
    attr_accessor :inserted
    # @return [FieldSimple] field simple
    attr_reader :field_simple
    # @return [Integer] id of paragraph (for comment)
    attr_accessor :paragraph_id
    # @return [MathParagraph] math paragraph
    attr_reader :math_paragraph
    # @return [Integer] id of text (for comment)
    attr_accessor :text_id

    def initialize(parent: nil)
      @number = 0
      @bookmark_start = []
      @bookmark_end = []
      @align = 'left'
      @spacing = Spacing.new
      @ind = Indents.new
      @character_style_array = []
      @horizontal_line = false
      @page_break = false
      @borders = Borders.new
      @keep_lines = false
      @contextual_spacing = false
      @page_numbering = false
      @keep_next = false
      @orphan_control = true
      @parent = parent
    end

    def initialize_copy(source)
      super
      @bookmark_start = source.bookmark_start.clone
      @bookmark_end = source.bookmark_end.clone
      @character_style_array = source.character_style_array.clone
      @spacing = source.spacing.clone
    end

    def nonempty_runs
      @character_style_array.select do |cur_run|
        if cur_run.is_a?(DocxParagraphRun) || cur_run.is_a?(ParagraphRun)
          !cur_run.empty?
        elsif cur_run.is_a?(DocxFormula) || cur_run.is_a?(StructuredDocumentTag)
          true
        end
      end
    end

    alias runs character_style_array

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
      ignored_attributes = %i[@number @parent]
      all_instance_variables = instance_variables
      significan_attribues = all_instance_variables - ignored_attributes
      significan_attribues.each do |current_attributes|
        return false unless instance_variable_get(current_attributes) == other.instance_variable_get(current_attributes)
      end
      true
    end

    def parse(node, par_number = 0, default_character = DocxParagraphRun.new, parent: nil)
      @parent ||= parent
      default_character_style = default_character.dup
      character_styles_array = []
      custom_character_style = default_character_style.dup
      custom_character_style.parent = self
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
          @field_simple = FieldSimple.new(parent: self).parse(node_child)
          @page_numbering = true if field_simple.page_numbering?
          character_styles_array += field_simple.runs
        when 'r'
          character_style = custom_character_style.dup
          node_child.xpath('w:instrText').each do |insrt_text|
            @page_numbering = true if insrt_text.text.include?('PAGE')
          end
          character_style.parse(node_child, char_number, parent: self)
          character_style.comments = comments.dup
          character_styles_array << character_style.dup
          character_styles_array.last.shape = character_style.shape unless character_style.shape.nil?
          char_number += 1
        when 'hyperlink'
          @hyperlink = Hyperlink.new(parent: self).parse(node_child)
          character_style = default_character_style.dup
          if !node_child.attribute('id').nil?
            character_style.link = Hyperlink.new(parent: character_style).parse(node_child)
          else
            character_style.link = node_child.attribute('anchor').value unless node_child.attribute('anchor').nil?
          end
          node_child.xpath('w:r').each do |r_tag|
            character_style.parse(r_tag, char_number, parent: self)
            character_styles_array << character_style.dup
            char_number += 1
          end
          node_child.xpath('w:fldSimple').each do |simple_field|
            hyperlink_field_simple = FieldSimple.new(parent: self).parse(simple_field)
            character_styles_array += hyperlink_field_simple.runs
          end
        when 'oMathPara'
          @math_paragraph = MathParagraph.new(parent: self).parse(node_child)
          character_styles_array << math_paragraph.math
        when 'commentRangeEnd'
          comments.each_with_index do |comment, index|
            if comment == node_child.attribute('id').value
              comments.delete_at(index)
              break
            end
          end
        when 'ins'
          @inserted = Inserted.new(parent: self).parse(node_child)
        when 'sdt'
          character_styles_array << StructuredDocumentTag.new(parent: self).parse(node_child)
        end
      end
      @number = par_number
      character_styles_array.last.text = character_styles_array.last.text.rstrip if character_styles_array.last.class == DocxParagraphRun
      @character_style_array = character_styles_array
      @parent = parent
      self
    end

    def parse_paragraph_style(node, default_char_style = DocxParagraphRun.new(parent: self))
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'pageBreakBefore'
          @page_break = true if node_child.attribute('val').nil? || node_child.attribute('val').value != 'false'
        when 'pBdr'
          @borders = ParagraphBorders.new(parent: self).parse(node_child)
        when 'keepLines'
          if node_child.attribute('val').nil?
            @keep_lines = true
          else
            @keep_lines = true unless node_child.attribute('val').value == 'false'
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
          @background_color.set_style(node_child.attribute('val').value.to_sym) unless node_child.attribute('val').nil?
        when 'pStyle'
          DocxParagraph.parse_paragraph_style_xml(node_child.attribute('val').value, self, default_char_style)
        when 'ind'
          @ind = DocumentStructure.default_paragraph_style.ind.dup.parse(node_child)
        when 'numPr'
          @numbering = NumberingProperties.new(parent: self).parse(node_child)
        when 'jc'
          @align = node_child.attribute('val').value.to_sym unless node_child.attribute('val').nil?
          @align = :justify if node_child.attribute('val').value == 'both'
        when 'spacing'
          @spacing.before = (node_child.attribute('before').value.to_f / 566.9).round(2) unless node_child.attribute('before').nil?
          @spacing.after = (node_child.attribute('after').value.to_f / 566.9).round(2) unless node_child.attribute('after').nil?
          @spacing.line_rule = node_child.attribute('lineRule').value.sub('atLeast', 'at_least').to_sym unless node_child.attribute('lineRule').nil?
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

    extend Gem::Deprecate
    deprecate :page_numbering, 'field_simple.page_numbering?', 2020, 1

    # @return [OoxmlParser::StructuredDocumentTag] Return first sdt element
    def sdt
      @character_style_array.each do |cur_element|
        return cur_element if cur_element.is_a?(StructuredDocumentTag)
      end
      nil
    end
    deprecate :sdt, 'nonempty_runs[i]', 2020, 1

    # @return [OoxmlParser::FrameProperties] Return frame properties
    def frame_properties
      paragraph_properties.frame_properties
    end
    deprecate :frame_properties, 'paragraph_properties.frame_properties', 2020, 1
  end
end
