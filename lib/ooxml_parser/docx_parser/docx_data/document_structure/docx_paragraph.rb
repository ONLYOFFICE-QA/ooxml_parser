# frozen_string_literal: true

# noinspection RubyTooManyInstanceVariablesInspection
require_relative 'docx_paragraph/bookmark_start'
require_relative 'docx_paragraph/bookmark_end'
require_relative 'docx_paragraph/comment_range_end'
require_relative 'docx_paragraph/comment_range_start'
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
  # Class for data of DocxParagraph
  class DocxParagraph < OOXMLDocumentObject
    include DocxParagraphHelper
    attr_accessor :number, :bookmark_start, :bookmark_end, :align, :spacing, :background_color, :ind, :numbering,
                  :character_style_array, :page_break, :borders, :keep_lines,
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
      @align = :left
      @spacing = Spacing.new
      @ind = Indents.new
      @character_style_array = []
      @page_break = false
      @borders = Borders.new
      @keep_lines = false
      @contextual_spacing = false
      @page_numbering = false
      @keep_next = false
      @orphan_control = true
      super
    end

    alias elements character_style_array

    # Constructor for copy of object
    # @param source [DocxParagraph] original object
    # @return [void]
    def initialize_copy(source)
      super
      @bookmark_start = source.bookmark_start.clone
      @bookmark_end = source.bookmark_end.clone
      @character_style_array = source.character_style_array.clone
      @spacing = source.spacing.clone
    end

    # @return [Array<OOXMLDocumentObject>] array of child objects that contains data
    def nonempty_runs
      @character_style_array.select do |cur_run|
        case cur_run
        when DocxParagraphRun, ParagraphRun
          !cur_run.empty?
        when DocxFormula, StructuredDocumentTag, BookmarkStart, BookmarkEnd, CommentRangeStart, CommentRangeEnd
          true
        end
      end
    end

    alias runs character_style_array

    # @return [True, false] if structure contain any user data
    def with_data?
      !nonempty_runs.empty? || paragraph_properties.section_properties
    end

    # Compare this object to other
    # @param other [Object] any other object
    # @return [True, False] result of comparision
    def ==(other)
      ignored_attributes = %i[@number @parent]
      all_instance_variables = instance_variables
      significan_attribues = all_instance_variables - ignored_attributes
      significan_attribues.each do |current_attributes|
        return false unless instance_variable_get(current_attributes) == other.instance_variable_get(current_attributes)
      end
      true
    end

    # Parse object
    # @param node [Nokogiri::XML:Node] node with DocxParagraph
    # @param par_number [Integer] number of paragraph
    # @param default_character [DocxParagraphRun] style for paragraph
    # @param parent [OOXMLDocumentObject] parent of run
    # @return [DocxParagraph] result of parse
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
          character_styles_array << BookmarkStart.new(parent: self).parse(node_child)
        when 'bookmarkEnd'
          character_styles_array << BookmarkEnd.new(parent: self).parse(node_child)
        when 'pPr'
          parse_paragraph_style(node_child, custom_character_style)
          @paragraph_properties = ParagraphProperties.new(parent: self).parse(node_child)
        when 'commentRangeStart'
          character_styles_array << CommentRangeStart.new(parent: self).parse(node_child)
        when 'commentRangeEnd'
          character_styles_array << CommentRangeEnd.new(parent: self).parse(node_child)
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
          character_styles_array.last.shape = character_style.shape if character_style.shape
          char_number += 1
        when 'hyperlink'
          @hyperlink = Hyperlink.new(parent: self).parse(node_child)
          character_style = default_character_style.dup
          character_style.parent = self
          if node_child.attribute('id')
            character_style.link = Hyperlink.new(parent: character_style).parse(node_child)
          elsif node_child.attribute('anchor')
            character_style.link = node_child.attribute('anchor').value
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
        when 'ins'
          @inserted = Inserted.new(parent: self).parse(node_child)
        when 'sdt'
          character_styles_array << StructuredDocumentTag.new(parent: self).parse(node_child)
        end
      end
      @number = par_number
      character_styles_array.last.text = character_styles_array.last.text.rstrip if character_styles_array.last.instance_of?(DocxParagraphRun)
      @character_style_array = character_styles_array
      @parent = parent
      self
    end

    # Parse style
    # @param node [Nokogiri::XML:Node] node with style
    # @param default_char_style [DocxParagraphRun] style for paragraph
    # @return [DocxParagraph] result of parse
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
          parse_paragraph_style_xml(node_child.attribute('val').value, default_char_style)
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

    # Parse style xml
    # @param id [String] id of style to parse
    # @param character_style [DocxParagraphRun] style to parse
    # @return [void]
    def parse_paragraph_style_xml(id, character_style)
      doc = parse_xml("#{OOXMLDocumentObject.path_to_folder}word/styles.xml")
      doc.search('//w:style').each do |style|
        next unless style.attribute('styleId').value == id

        style.xpath('w:pPr').each do |p_pr|
          parse_paragraph_style(p_pr, character_style)
          @style = StyleParametres.new(parent: self).parse(style)
        end
        style.xpath('w:rPr').each do |r_pr|
          character_style.parse_properties(r_pr)
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
