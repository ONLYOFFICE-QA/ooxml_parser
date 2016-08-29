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

    def parse(p_tag, par_number = 0, default_character = DocxParagraphRun.new, parent: nil)
      @parent = parent
      default_character_style = default_character.copy
      character_styles_array = []
      custom_character_style = default_character_style.copy
      char_number = 0
      comments = []
      p_tag.xpath('w:bookmarkStart').each do |bookmark_start|
        @bookmark_start << Bookmark.new(bookmark_start.attribute('id').value, bookmark_start.attribute('name').value)
      end
      p_tag.xpath('w:bookmarkEnd').each do |bookmark_end|
        @bookmark_end << Bookmark.new(bookmark_end.attribute('id').value)
      end
      p_tag.xpath('*').each do |p_element|
        if p_element.name == 'pPr'
          p_props = p_tag.xpath('w:pPr')
          DocxParagraph.parse_paragraph_style(p_props, self, custom_character_style, parent: parent)
          p_tag.xpath('w:pict').each do |pict|
            pict.xpath('v:rect').each do
              @horizontal_line = true
            end
          end
          @paragraph_properties = ParagraphProperties.parse(p_element)
        elsif p_element.name == 'commentRangeStart'
          comments << p_element.attribute('id').value
        elsif p_element.name == 'fldSimple'
          instruction = p_element.attribute('instr').to_s
          @page_numbering = true if instruction.include?('PAGE')
          p_element.xpath('w:r').each do |r_tag|
            character_style = default_character_style.copy
            character_style.parse(r_tag, char_number, parent: parent)
            character_style.page_number = @page_numbering
            character_style.instruction = instruction
            character_styles_array << character_style.copy
            char_number += 1
          end
        elsif p_element.name == 'r'
          character_style = custom_character_style.copy
          p_element.xpath('w:instrText').each do |insrt_text|
            @page_numbering = true if insrt_text.text.include?('PAGE')
          end
          character_style.parse(p_element, char_number, parent: self)
          character_style.comments = comments.dup
          character_styles_array << character_style.copy
          unless character_style.shape.nil?
            character_styles_array.last.shape = character_style.shape
          end
          char_number += 1
        elsif p_element.name == 'hyperlink'
          character_style = default_character_style.copy
          if !p_element.attribute('id').nil?
            character_style.link = Hyperlink.parse(p_element)
          else
            unless p_element.attribute('anchor').nil?
              character_style.link = p_element.attribute('anchor').value
            end
          end
          p_element.xpath('w:r').each do |r_tag|
            character_style.parse(r_tag, char_number, parent: parent)
            character_styles_array << character_style.copy
            char_number += 1
          end
          p_element.xpath('w:fldSimple').each do |simple_field|
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
        elsif p_element.name == 'oMathPara'
          p_element.xpath('m:oMath').each do |o_math|
            character_styles_array << DocxFormula.parse(o_math)
          end
        elsif p_element.name == 'commentRangeEnd'
          comments.each_with_index do |comment, index|
            if comment == p_element.attribute('id').value
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

    def self.parse_paragraph_style(paragraph_pr_tag,
                                   paragraph_style = DocxParagraph.new,
                                   default_char_style = DocxParagraphRun.new,
                                   parent: nil)
      paragraph_pr_tag.xpath('w:tabs').each do |tabs_node|
        tabs_node.xpath('w:tab').each { |tab_node| paragraph_style.tabs << ParagraphTab.new(tab_node.attribute('val').value.to_sym, (tab_node.attribute('pos').value.to_f / 566.9).round(2)) }
      end
      paragraph_pr_tag.xpath('w:pageBreakBefore').each do |page_break_before|
        if page_break_before.attribute('val').nil? || page_break_before.attribute('val').value != 'false'
          paragraph_style.page_break = true
        end
      end
      paragraph_pr_tag.xpath('w:pBdr').each do |paragraph_br|
        paragraph_style.borders = Borders.new
        paragraph_br.xpath('w:bottom').each do |bottom|
          paragraph_style.borders.bottom = BordersProperties.parse(bottom)
        end
        paragraph_br.xpath('w:left').each do |left|
          paragraph_style.borders.left = BordersProperties.parse(left)
        end
        paragraph_br.xpath('w:top').each do |top|
          paragraph_style.borders.top = BordersProperties.parse(top)
        end
        paragraph_br.xpath('w:right').each do |right|
          paragraph_style.borders.right = BordersProperties.parse(right)
        end
        paragraph_br.xpath('w:between').each do |between|
          paragraph_style.borders.between = BordersProperties.parse(between)
        end
        paragraph_br.xpath('w:bar').each do |bar|
          paragraph_style.borders.bar = BordersProperties.parse(bar)
        end
      end
      paragraph_pr_tag.xpath('w:keepLines').each do |keep_lines|
        if keep_lines.attribute('val').nil?
          paragraph_style.keep_lines = true
        else
          unless keep_lines.attribute('val').value == 'false'
            paragraph_style.keep_lines = true
          end
        end
      end
      paragraph_pr_tag.xpath('w:widowControl').each do |widow_control_node|
        paragraph_style.orphan_control = OOXMLDocumentObject.option_enabled?(widow_control_node)
      end
      paragraph_pr_tag.xpath('w:keepNext').each do |_|
        paragraph_style.keep_next = true
      end
      paragraph_style.contextual_spacing = true unless paragraph_pr_tag.xpath('w:contextualSpacing').empty?
      paragraph_pr_tag.xpath('w:shd').each do |shd|
        background_color_string = shd.attribute('fill').value
        paragraph_style.background_color = Color.from_int16(background_color_string)
        unless shd.attribute('val').nil?
          paragraph_style.background_color.set_style(shd.attribute('val').value)
        end
      end
      paragraph_pr_tag.xpath('w:pStyle').each do |p_style|
        parse_paragraph_style_xml(p_style.attribute('val').value, paragraph_style, default_char_style)
      end
      paragraph_pr_tag.xpath('w:ind').each do |ind|
        paragraph_style.ind = Indents.new(parent: paragraph_style).parse(ind)
      end
      paragraph_pr_tag.xpath('w:kinoku').each do
        paragraph_style.kinoku = true
      end
      paragraph_pr_tag.xpath('w:framePr').each do |frame_pr_node|
        paragraph_style.frame_properties = FrameProperties.parse(frame_pr_node)
      end
      paragraph_pr_tag.xpath('w:numPr').each do |num_pr|
        paragraph_style.numbering = NumberingProperties.parse(num_pr, paragraph_style)
      end
      paragraph_pr_tag.xpath('w:jc').each do |jc|
        paragraph_style.align = jc.attribute('val').value.to_sym unless jc.attribute('val').nil?
        paragraph_style.align = :justify if jc.attribute('val').value == 'both'
      end
      paragraph_pr_tag.xpath('w:framePr').each do |frame_pr_node|
        paragraph_style.frame_properties = FrameProperties.parse(frame_pr_node)
      end
      paragraph_pr_tag.xpath('w:spacing').each do |spacing|
        unless spacing.attribute('before').nil?
          paragraph_style.spacing.before = (spacing.attribute('before').value.to_f / 566.9).round(2)
        end
        unless spacing.attribute('after').nil?
          paragraph_style.spacing.after = (spacing.attribute('after').value.to_f / 566.9).round(2)
        end
        unless spacing.attribute('lineRule').nil?
          paragraph_style.spacing.line_rule = spacing.attribute('lineRule').value.sub('atLeast', 'at_least').to_sym
        end
        unless spacing.attribute('line').nil?
          paragraph_style.spacing.line = (paragraph_style.spacing.line_rule == :auto ? (spacing.attribute('line').value.to_f / 240.0).round(2) : (spacing.attribute('line').value.to_f / 566.9).round(2))
        end
      end
      paragraph_pr_tag.xpath('w:sectPr').each do |sect_pr|
        paragraph_style.sector_properties = PageProperties.parse(sect_pr, paragraph_style, default_char_style, parent: parent)
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
