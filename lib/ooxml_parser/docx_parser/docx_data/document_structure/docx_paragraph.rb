# noinspection RubyTooManyInstanceVariablesInspection
require_relative 'docx_paragraph/docx_paragraph_run'
require_relative 'docx_paragraph/indents'
require_relative 'docx_paragraph/frame_properties'
require_relative 'docx_paragraph/docx_formula'
require_relative 'docx_paragraph/paragraph_style'
module OoxmlParser
  class DocxParagraph < OOXMLDocumentObject
    attr_accessor :number, :bookmark_start, :bookmark_end, :align, :spacing, :background_color, :ind, :numbering,
                  :character_style_array, :horizontal_line, :page_break, :kinoku, :borders, :keep_lines,
                  :contextual_spacing, :sector_properties, :page_numbering, :section_break, :style, :keep_next,
                  :orphan_control, :tabs, :frame_properties
    # @return [ParagraphProperties] Properties of current paragraph
    attr_accessor :paragraph_properties

    def initialize
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
      @tabs = []
      @frame_properties = nil
    end

    def copy
      paragraph = DocxParagraph.new
      paragraph.number = number
      paragraph.bookmark_start = @bookmark_start.dup
      paragraph.bookmark_end = @bookmark_end.dup
      paragraph.align = @align
      paragraph.spacing = @spacing.copy
      paragraph.background_color = @background_color
      paragraph.ind = @ind.copy
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
      paragraph.tabs = @tabs.dup
      paragraph.frame_properties = @frame_properties
      paragraph.paragraph_properties = @paragraph_properties
      paragraph
    end

    def nonempty_runs
      @character_style_array.select do |cur_run|
        if cur_run.is_a?(DocxParagraphRun)
          (!cur_run.text.empty? || !cur_run.alternate_content.nil? || !cur_run.drawing.nil?)
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
      character_style_array.each do |current_run|
        character_style_array.delete(current_run) if current_run.text.empty?
      end
      other.character_style_array.each do |current_run|
        other.character_style_array.delete(current_run) if current_run.text.empty?
      end
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
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'bookmarkStart'
          @bookmark_start << Bookmark.new(node_child.attribute('id').value, node_child.attribute('name').value)
        when 'bookmarkEnd'
          @bookmark_end << Bookmark.new(node_child.attribute('id').value)
        when 'pPr'
          DocxParagraph.parse_paragraph_style(node_child, self, custom_character_style, parent: parent)
          node.xpath('w:pict').each do |pict|
            pict.xpath('v:rect').each do
              @horizontal_line = true
            end
          end
          @paragraph_properties = ParagraphProperties.parse(node_child)
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
            character_style.link = Hyperlink.parse(node_child)
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
            character_styles_array << DocxFormula.parse(o_math)
          end
        when 'commentRangeEnd'
          comments.each_with_index do |comment, index|
            if comment == node_child.attribute('id').value
              comments.delete_at(index)
              break
            end
          end
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

    def self.parse_paragraph_style(node,
                                   paragraph_style = DocxParagraph.new,
                                   default_char_style = DocxParagraphRun.new,
                                   parent: nil)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'tabs'
          node_child.xpath('w:tab').each { |tab_node| paragraph_style.tabs << ParagraphTab.new(tab_node.attribute('val').value.to_sym, (tab_node.attribute('pos').value.to_f / 566.9).round(2)) }
        when 'pageBreakBefore'
          if node_child.attribute('val').nil? || node_child.attribute('val').value != 'false'
            paragraph_style.page_break = true
          end
        when 'pBdr'
          paragraph_style.borders = ParagraphBorders.parse(node_child)
        when 'keepLines'
          if node_child.attribute('val').nil?
            paragraph_style.keep_lines = true
          else
            unless node_child.attribute('val').value == 'false'
              paragraph_style.keep_lines = true
            end
          end
        when 'widowControl'
          paragraph_style.orphan_control = OOXMLDocumentObject.option_enabled?(node_child)
        when 'keepNext'
          paragraph_style.keep_next = true
        when 'contextualSpacing'
          paragraph_style.contextual_spacing = true
        when 'shd'
          background_color_string = node_child.attribute('fill').value
          paragraph_style.background_color = Color.from_int16(background_color_string)
          unless node_child.attribute('val').nil?
            paragraph_style.background_color.set_style(node_child.attribute('val').value)
          end
        when 'pStyle'
          parse_paragraph_style_xml(node_child.attribute('val').value, paragraph_style, default_char_style)
        when 'ind'
          paragraph_style.ind = Indents.new(parent: paragraph_style).parse(node_child)
        when 'kinoku'
          paragraph_style.kinoku = true
        when 'framePr'
          paragraph_style.frame_properties = FrameProperties.parse(node_child)
        when 'numPr'
          paragraph_style.numbering = NumberingProperties.parse(node_child, paragraph_style)
        when 'jc'
          paragraph_style.align = node_child.attribute('val').value.to_sym unless node_child.attribute('val').nil?
          paragraph_style.align = :justify if node_child.attribute('val').value == 'both'
        when 'spacing'
          unless node_child.attribute('before').nil?
            paragraph_style.spacing.before = (node_child.attribute('before').value.to_f / 566.9).round(2)
          end
          unless node_child.attribute('after').nil?
            paragraph_style.spacing.after = (node_child.attribute('after').value.to_f / 566.9).round(2)
          end
          unless node_child.attribute('lineRule').nil?
            paragraph_style.spacing.line_rule = node_child.attribute('lineRule').value.sub('atLeast', 'at_least').to_sym
          end
          unless node_child.attribute('line').nil?
            paragraph_style.spacing.line = (paragraph_style.spacing.line_rule == :auto ? (node_child.attribute('line').value.to_f / 240.0).round(2) : (node_child.attribute('line').value.to_f / 566.9).round(2))
          end
        when 'sectPr'
          paragraph_style.sector_properties = PageProperties.new(parent: paragraph_style).parse(node_child, paragraph_style, default_char_style)
          paragraph_style.section_break = case paragraph_style.sector_properties.type
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
      paragraph_style.parent = parent
      paragraph_style
    end

    def self.parse_paragraph_style_xml(id, paragraph_style, character_style)
      doc = Nokogiri::XML(File.open(OOXMLDocumentObject.path_to_folder + 'word/styles.xml'))
      doc.search('//w:style').each do |style|
        next unless style.attribute('styleId').value == id
        style.xpath('w:pPr').each do |p_pr|
          parse_paragraph_style(p_pr, paragraph_style, character_style)
          paragraph_style.style = StyleParametres.parse(style)
        end
        style.xpath('w:rPr').each do |r_pr|
          character_style.parse_properties(r_pr, DocumentStructure.default_run_style)
        end
        break
      end
    end
  end
end
