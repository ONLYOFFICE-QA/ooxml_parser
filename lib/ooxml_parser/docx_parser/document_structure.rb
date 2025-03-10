# frozen_string_literal: true

require_relative 'document_structure/comments'
require_relative 'document_structure/comments_extended'
require_relative 'document_structure/default_style_helper'
require_relative 'document_structure/docx_paragraph'
require_relative 'document_structure/document_background'
require_relative 'document_structure/document_properties'
require_relative 'document_structure/document_structure_helpers'
require_relative 'document_structure/document_style'
require_relative 'document_structure/header_footer'
require_relative 'document_structure/numbering'
require_relative 'document_structure/page_properties/page_properties'
require_relative 'document_structure/document_settings'
require_relative 'document_structure/styles'
module OoxmlParser
  # Basic class for DocumentStructure
  class DocumentStructure < CommonDocumentStructure
    include DefaultStyleHelper
    include DocumentStyleHelper
    include DocumentStructureHelpers
    # @return [Array<OOXMLDocumentObject>] list of elements
    attr_accessor :elements
    # @return [PageProperties] properties of document
    attr_accessor :page_properties
    # @return [Note] notes of document
    attr_accessor :notes
    # @return [DocumentBackground] background of document
    attr_accessor :background
    # @return [DocumentProperties] properties of document
    attr_accessor :document_properties
    # @return [Comments] comment of document
    attr_accessor :comments
    # @return [Numbering] store numbering data
    attr_accessor :numbering
    # @return [Styles] styles of document
    attr_accessor :styles
    # @return [PresentationTheme] theme of docx
    attr_accessor :theme
    # @return [Relationships] relationships
    attr_accessor :relationships
    # @return [DocumentSettings] settings
    attr_accessor :settings
    # @return [CommentsDocument] comments of whole document
    attr_accessor :comments_document
    # @return [CommentsExtended] extended comments
    attr_accessor :comments_extended
    # @return [DocxParagraph] default paragraph style
    attr_accessor :default_paragraph_style
    # @return [DocxParagraphRun] default run style
    attr_accessor :default_run_style
    # @return [DocxParagraph] default table paragraph style
    attr_accessor :default_table_paragraph_style
    # @return [DocxParagraphRun] default table run style
    attr_accessor :default_table_run_style

    def initialize(params = {})
      @elements = []
      @notes = []
      @document_properties = DocumentProperties.new
      @comments = []
      super
    end

    alias theme_colors theme

    # Compare this object to other
    # @param other [Object] any other object
    # @return [True, False] result of comparision
    def ==(other)
      @elements == other.elements &&
        @page_properties == other.page_properties &&
        @notes == other.notes &&
        @background == other.background &&
        @document_properties == other.document_properties &&
        @comments == other.comments
    end

    # Get element by it's type
    # @param location [Symbol] location of object
    # @param type [Symbol] type of object
    # @return [OOXMLDocumentObject]
    def element_by_description(location: :canvas, type: :docx_paragraph)
      case location
      when :canvas
        case type
        when :table
          elements[1].rows[0].cells[0].elements
        when :docx_paragraph, :simple, :paragraph
          elements
        when :shape
          elements[0].nonempty_runs.first.alternate_content.office2007_content.data.text_box
        else
          raise 'Wrong location(Need One of ":table", ":paragraph", ":shape")'
        end
      when :footer
        case type
        when :table
          note_by_description(:footer1).elements[0].rows[0].cells[0].elements
        when :docx_paragraph, :simple, :paragraph
          note_by_description(:footer1).elements
        when :shape
          note_by_description(:footer1).elements[0].nonempty_runs.first.alternate_content.office2007_content.data.text_box
        else
          raise 'Wrong location(Need One of ":table", ":simple", ":shape")'
        end
      when :header
        case type
        when :table
          note_by_description(:header1).elements[0].rows[0].cells[0].elements
        when :docx_paragraph, :simple, :paragraph
          note_by_description(:header1).elements
        when :shape
          note_by_description(:header1).elements[0].nonempty_runs.first.alternate_content.office2007_content.data.text_box
        else
          raise 'Wrong location(Need One of ":table", ":simple", ":shape")'
        end
      when :comment
        comments[0].paragraphs
      else
        raise 'Wrong global location(Need One of ":canvas", ":footer", ":header", ":comment")'
      end
    end

    # Get note by it's description
    # @param type [Symbol] note type
    # @return [Note]
    def note_by_description(type)
      notes.each do |note|
        return note if note.type.to_sym == type
      end
      raise 'There isn\'t this type of the note'
    end

    # Detect numbering type
    # @param location [Symbol] location of object
    # @param type [Symbol] type of object
    # @param paragraph_number [Integer] number of object
    # @return [Array<String,String>] type of numbering
    def recognize_numbering(location: :canvas, type: :simple, paragraph_number: 0)
      elements = element_by_description(location: location, type: type)
      lvl_text = elements[paragraph_number].numbering.abstruct_numbering.level_list[0].text.value
      num_format = elements[paragraph_number].numbering.abstruct_numbering.level_list[0].numbering_format.value
      [num_format, lvl_text]
    end

    # Return outline type
    # @param location [Symbol] location of object
    # @param type [Symbol] type of object
    # @param levels_count [Integer] count of levels to detect
    # @return [Array<String,String>] type of outline
    def outline(location: :canvas, type: :simple, levels_count: 1)
      elements = element_by_description(location: location, type: type)
      set = []
      levels_count.times do |col|
        set[0] = elements[col].numbering.abstruct_numbering.level_list[col].numbering_format.value
        set[1] = elements[col].numbering.abstruct_numbering.level_list[col].text.value
      end
      set
    end

    # @return [Array<DocumentStyle>] style of documents
    def document_styles
      styles.styles
    end

    # Parse docx file
    # @return [DocumentStructure] parsed structure
    def parse
      @content_types = ContentTypes.new(parent: self).parse
      @root_subfolder = 'word/'
      @comments = []
      @default_paragraph_style = DocxParagraph.new
      @default_run_style = DocxParagraphRun.new(parent: self)
      @theme = PresentationTheme.new(parent: self).parse('word/theme/theme1.xml')
      @relationships = Relationships.new(parent: self).parse_file("#{root_object.unpacked_folder}word/_rels/document.xml.rels")
      parse_styles
      number = 0
      root_object.add_to_xmls_stack('word/document.xml')
      doc = parse_xml(root_object.current_xml)
      doc.search('//w:document').each do |document|
        document.xpath('w:background').each do |background|
          @background = DocumentBackground.new(parent: self).parse(background)
        end
        document.xpath('w:body').each do |body|
          body.xpath('*').each do |element|
            case element.name
            when 'p'
              child = element.child
              unless child.nil? && @elements.last.instance_of?(Table)
                paragraph_style = default_paragraph_style.dup.parse(element, number, default_run_style, parent: self)
                number += 1
                @elements << paragraph_style.dup
              end
            when 'tbl'
              table = Table.new(parent: self).parse(element,
                                                    number,
                                                    TableProperties.new)
              number += 1
              @elements << table
            when 'sdt'
              @elements << StructuredDocumentTag.new(parent: self).parse(element)
            end
          end
          body.xpath('w:sectPr').each do |sect_pr|
            @page_properties = PageProperties.new(parent: self).parse(sect_pr,
                                                                      default_paragraph_style,
                                                                      default_run_style)
            @notes = page_properties.notes # keep copy of notes to compatibility with previous docx models
          end
        end
      end
      root_object.xmls_stack.pop
      @document_properties = DocumentProperties.new(parent: self).parse
      @comments = Comments.new(parent: self).parse
      @comments_extended = CommentsExtended.new(parent: self).parse
      @comments_document = Comments.new(parent: self,
                                        file: "#{root_object.unpacked_folder}word/#{relationships.target_by_type('commentsDocument').first}")
                                   .parse
      @settings = DocumentSettings.new(parent: self).parse
      self
    end
  end
end
