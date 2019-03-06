# noinspection RubyInstanceMethodNamingConvention
require_relative 'document_structure/comments'
require_relative 'document_structure/comments_extended'
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
  class DocumentStructure < CommonDocumentStructure
    include DocumentStyleHelper
    include DocumentStructureHelpers
    attr_accessor :elements, :page_properties, :notes, :background, :document_properties, :comments

    # @return [Numbering] store numbering data
    attr_accessor :numbering
    # @return [Styles] styles of document
    attr_accessor :styles
    attr_accessor :theme
    # @return [Relationships] relationships
    attr_accessor :relationships
    # @return [DocumentSettings] settings
    attr_accessor :settings
    # @return [CommentsDocument] comments of whole document
    attr_accessor :comments_document
    # @return [CommentsExtended] extended comments
    attr_accessor :comments_extended

    def initialize
      @elements = []
      @notes = []
      @document_properties = DocumentProperties.new
      @comments = []
      super
    end

    alias theme_colors theme

    def ==(other)
      @elements == other.elements &&
        @page_properties == other.page_properties &&
        @notes == other.notes &&
        @background == other.background &&
        @document_properties == other.document_properties &&
        @comments == other.comments
    end

    def difference(other)
      Hash.object_to_hash(self).diff(Hash.object_to_hash(other))
    end

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
          note_by_description(:footer1).elements[1].rows[0].cells[0].elements
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
          note_by_description(:header1).elements[1].rows[0].cells[0].elements
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

    def note_by_description(type)
      notes.each do |note|
        return note if note.type.to_sym == type
      end
      raise 'There isn\'t this type of the note'
    end

    def recognize_numbering(location: :canvas, type: :simple, paragraph_number: 0)
      elements = element_by_description(location: location, type: type)
      lvl_text = elements[paragraph_number].numbering.abstruct_numbering.level_list[0].text.value
      num_format = elements[paragraph_number].numbering.abstruct_numbering.level_list[0].numbering_format.value
      [num_format, lvl_text]
    end

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

    def self.parse
      doc_structure = DocumentStructure.new
      doc_structure.content_types = ContentTypes.new(parent: doc_structure).parse
      OOXMLDocumentObject.root_subfolder = 'word/'
      OOXMLDocumentObject.xmls_stack = []
      @comments = []
      DocumentStructure.default_paragraph_style = DocxParagraph.new
      DocumentStructure.default_run_style = DocxParagraphRun.new(parent: doc_structure)
      doc_structure.theme = PresentationTheme.parse('word/theme/theme1.xml')
      doc_structure.relationships = Relationships.new(parent: self).parse_file("#{OOXMLDocumentObject.path_to_folder}word/_rels/document.xml.rels")
      OOXMLDocumentObject.add_to_xmls_stack('word/styles.xml')
      doc = doc_structure.parse_xml(OOXMLDocumentObject.current_xml)
      # TODO: Remove this old way parsing in favor of doc_structure.styles.document_defaults
      doc.search('//w:docDefaults').each do |doc_defaults|
        doc_defaults.xpath('w:pPrDefault').each do |p_pr_defaults|
          DocumentStructure.default_paragraph_style = DocxParagraph.new(parent: doc_structure).parse(p_pr_defaults, 0)
        end
        doc_defaults.xpath('w:rPrDefault').each do |r_pr_defaults|
          r_pr_defaults.xpath('w:rPr').each do |r_pr|
            DocumentStructure.default_run_style = DocxParagraphRun.new(parent: doc_structure).parse_properties(r_pr)
          end
        end
      end
      doc_structure.parse_default_style
      doc_structure.numbering = Numbering.new(parent: doc_structure).parse
      doc_structure.styles = Styles.new(parent: doc_structure).parse
      number = 0
      OOXMLDocumentObject.add_to_xmls_stack('word/document.xml')
      doc = doc_structure.parse_xml(OOXMLDocumentObject.current_xml)
      doc.search('//w:document').each do |document|
        document.xpath('w:background').each do |background|
          doc_structure.background = DocumentBackground.new(parent: doc_structure).parse(background)
        end
        document.xpath('w:body').each do |body|
          body.xpath('*').each do |element|
            case element.name
            when 'p'
              child = element.child
              unless child.nil? && doc_structure.elements.last.class == Table
                paragraph_style = DocumentStructure.default_paragraph_style.dup.parse(element, number, DocumentStructure.default_run_style, parent: doc_structure)
                number += 1
                doc_structure.elements << paragraph_style.dup
              end
            when 'tbl'
              table = Table.new(parent: doc_structure).parse(element,
                                                             number,
                                                             TableProperties.new)
              number += 1
              doc_structure.elements << table
            when 'sdt'
              doc_structure.elements << StructuredDocumentTag.new(parent: doc_structure).parse(element)
            end
          end
          body.xpath('w:sectPr').each do |sect_pr|
            doc_structure.page_properties = PageProperties.new(parent: doc_structure).parse(sect_pr,
                                                                                            DocumentStructure.default_paragraph_style,
                                                                                            DocumentStructure.default_run_style)
            doc_structure.notes = doc_structure.page_properties.notes # keep copy of notes to compatibility with previous docx models
          end
        end
      end
      OOXMLDocumentObject.xmls_stack.pop
      doc_structure.document_properties = DocumentProperties.new(parent: doc_structure).parse
      doc_structure.comments = Comments.new(parent: doc_structure).parse
      doc_structure.comments_extended = CommentsExtended.new(parent: doc_structure).parse
      doc_structure.comments_document = Comments.new(parent: doc_structure,
                                                     file: "#{OOXMLDocumentObject.path_to_folder}word/#{doc_structure.relationships.target_by_type('commentsDocument').first}")
                                                .parse
      doc_structure.settings = DocumentSettings.new(parent: doc_structure).parse
      doc_structure
    end

    def parse_default_style
      doc = parse_xml(OOXMLDocumentObject.path_to_folder + 'word/styles.xml')
      doc.search('//w:style').each do |style|
        next if style.attribute('default').nil?

        if (style.attribute('default').value == '1' || style.attribute('default').value == 'on' || style.attribute('default').value == 'true') && style.attribute('type').value == 'paragraph'
          style.xpath('w:pPr').each do |paragraph_pr_tag|
            DocumentStructure.default_paragraph_style = DocxParagraph.new.parse_paragraph_style(paragraph_pr_tag, DocumentStructure.default_run_style)
          end
          style.xpath('w:rPr').each do |character_pr_tag|
            DocumentStructure.default_run_style.parse_properties(character_pr_tag, DocumentStructure.default_run_style)
          end
        elsif (style.attribute('default').value == '1' || style.attribute('default').value == 'on' || style.attribute('default').value == 'true') && style.attribute('type').value == 'character'
          style.xpath('w:rPr').each do |character_pr_tag|
            DocumentStructure.default_run_style.parse_properties(character_pr_tag, DocumentStructure.default_run_style)
          end
        end
      end
      DocumentStructure.default_table_paragraph_style = DocumentStructure.default_paragraph_style.dup
      DocumentStructure.default_table_paragraph_style.spacing = Spacing.new(0, 0, 1, :auto)
      DocumentStructure.default_table_run_style = DocumentStructure.default_run_style.dup
      doc.search('//w:style').each do |style|
        next if style.attribute('default').nil?
        next unless (style.attribute('default').value == '1' || style.attribute('default').value == 'on' || style.attribute('default').value == 'true') && style.attribute('type').value == 'table'

        style.xpath('w:rPr').each do |table_character_pr_tag|
          DocumentStructure.default_table_run_style.parse_properties(table_character_pr_tag, DocumentStructure.default_run_style)
        end
      end
    end

    class << self
      attr_accessor :default_table_run_style
      attr_accessor :default_table_paragraph_style
      attr_accessor :default_paragraph_style
      attr_accessor :default_run_style
    end
  end
end
