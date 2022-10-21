# frozen_string_literal: true

require_relative 'docx_paragraph_run/docx_paragraph_run_helpers'
require_relative 'docx_paragraph_run/instruction_text'
require_relative 'docx_paragraph_run/object'
require_relative 'docx_paragraph_run/text_outline'
require_relative 'docx_paragraph_run/text_fill'
require_relative 'docx_paragraph_run/shape'

module OoxmlParser
  # Class for working with DocxParagraphRun
  class DocxParagraphRun < OOXMLDocumentObject
    include DocxParagraphRunHelpers
    attr_accessor :number, :font, :vertical_align, :size, :font_color, :font_style, :text, :drawings,
                  :link, :highlight, :effect, :caps, :w,
                  :position, :em, :spacing, :break, :touch, :shape, :footnote, :endnote, :fld_char, :style,
                  :comments, :alternate_content, :page_number, :text_outline, :text_fill
    # @return [InstructionText] text of instruction
    #   See ECMA-376, 17.16.23 instrText (Field Code)
    attr_reader :instruction_text
    # @return [RunProperties] properties of run
    attr_accessor :run_properties
    # @return [RunObject] object of run
    attr_accessor :object
    # @return [Shade] shade properties
    attr_accessor :shade

    def initialize(parent: nil)
      @number = 0
      @font = 'Arial'
      @vertical_align = :baseline
      @size = 11
      @font_color = Color.new
      @font_style = FontStyle.new
      @text = ''
      @drawings = []
      @w = false
      @position = 0.0
      @spacing = 0.0
      @break = false
      @comments = []
      @page_number = false
      super
    end

    # Constructor for copy of object
    # @param source [DocxParagraphRun] original object
    # @return [void]
    def initialize_copy(source)
      super
      @drawings = source.drawings.clone
      @comments = source.comments.clone
    end

    # @return [Array, nil] Drawings of Run
    def drawing
      # TODO: Rewrite all tests without this methos
      @drawings.empty? ? nil : drawings.first
    end

    # @return [True, False] is current run empty
    def empty?
      text.empty? &&
        !alternate_content &&
        !drawing &&
        !object &&
        !shape &&
        !footnote &&
        !endnote &&
        !@break
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
    # @param r_tag [Nokogiri::XML:Node] node with DocxParagraphRun
    # @param char_number [Integer] number of run
    # @param parent [OOXMLDocumentObject] parent of run
    # @return [void]
    def parse(r_tag, char_number, parent: nil)
      @parent = parent
      r_tag.xpath('*').each do |node_child|
        case node_child.name
        when 'rPr'
          parse_properties(node_child)
          @run_properties = RunProperties.new(parent: self).parse(node_child)
        when 'instrText'
          @instruction_text = InstructionText.new(parent: self).parse(node_child)
          if @instruction_text.hyperlink?
            @link = @instruction_text.to_hyperlink
          elsif @instruction_text.page_number?
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
          @drawings << DocxDrawing.new(parent: self).parse(node_child)
        when 'AlternateContent'
          @alternate_content = AlternateContent.new(parent: self).parse(node_child)
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
          @footnote = HeaderFooter.new(parent: self).parse(node_child)
        when 'endnoteReference'
          @endnote = HeaderFooter.new(parent: self).parse(node_child)
        when 'pict'
          node_child.xpath('*').each do |pict_node_child|
            case pict_node_child.name
            when 'shape'
              @shape = Shape.new(parent: self).parse(pict_node_child, :shape)
            when 'rect'
              @shape = Shape.new(parent: self).parse(pict_node_child, :rectangle)
            end
          end
        when 'object'
          @object = RunObject.new(parent: self).parse(node_child)
        end
      end
      @number = char_number
    end
  end
end
