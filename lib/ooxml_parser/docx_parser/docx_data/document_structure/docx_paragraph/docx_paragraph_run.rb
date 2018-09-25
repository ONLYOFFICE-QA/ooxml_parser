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
                  :link, :highlight, :effect, :caps, :w,
                  :position, :em, :spacing, :break, :touch, :shape, :footnote, :endnote, :fld_char, :style,
                  :comments, :alternate_content, :page_number, :text_outline, :text_fill
    # @return [String] type of instruction used for upper level of run
    # http://officeopenxml.com/WPfieldInstructions.php
    attr_accessor :instruction
    # @return [RunProperties] properties of run
    attr_accessor :run_properties
    # @return [RunObject] object of run
    attr_accessor :object

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
      @parent = parent
    end

    def initialize_copy(source)
      super
      @drawings = source.drawings.clone
      @comments = source.comments.clone
    end

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

    def ==(other)
      ignored_attributes = %i[@number @parent]
      all_instance_variables = instance_variables
      significan_attribues = all_instance_variables - ignored_attributes
      significan_attribues.each do |current_attributes|
        return false unless instance_variable_get(current_attributes) == other.instance_variable_get(current_attributes)
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
            when 'oval'
              @shape = Shape.new(parent: self).parse(pict_node_child, :oval)
            end
          end
        when 'object'
          @object = RunObject.new(parent: self).parse(node_child)
        end
      end
      @number = char_number
    end

    def self.parse_font_by_theme(theme)
      theme_file = "#{OOXMLDocumentObject.path_to_folder}/#{OOXMLDocumentObject.root_subfolder}/theme/theme1.xml"
      return nil unless File.exist?(theme_file)

      doc = Nokogiri::XML(File.open(theme_file))
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
