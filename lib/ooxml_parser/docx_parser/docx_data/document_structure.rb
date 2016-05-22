# noinspection RubyInstanceMethodNamingConvention
require_relative 'document_structure/comment'
require_relative 'document_structure/docx_paragraph'
require_relative 'document_structure/document_background'
require_relative 'document_structure/document_properties'
require_relative 'document_structure/document_style'
require_relative 'document_structure/header_footer'
require_relative 'document_structure/page_properties/page_properties'
module OoxmlParser
  class DocumentStructure < CommonDocumentStructure
    attr_accessor :elements, :page_properties, :notes, :background, :document_properties, :comments

    # @return [Array, DocumentStyle] array of document styles in current document
    attr_accessor :document_styles

    def initialize(elements = [], page_properties = nil, notes = [], background = nil, document_properties = DocumentProperties.new, comments = [])
      @elements = elements
      @page_properties = page_properties
      @notes = notes
      @background = background
      @document_properties = document_properties
      @comments = comments
      @document_styles = []
    end

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
      lvl_text = elements[paragraph_number].numbering.numbering_properties.ilvls[0].level_text
      num_format = elements[paragraph_number].numbering.numbering_properties.ilvls[0].num_format
      [num_format, lvl_text]
    end

    def outline(location: :canvas, type: :simple, levels_count: 1)
      elements = element_by_description(location: location, type: type)
      set = []
      levels_count.times do |col|
        set[0] = elements[col].numbering.numbering_properties.ilvls[col].num_format
        set[1] = elements[col].numbering.numbering_properties.ilvls[col].level_text
      end
      set
    end

    # Return document style by its name
    # @param name [String] name of style
    # @return [DocumentStyle, nil]
    def document_style_by_name(name)
      @document_styles.each do |style|
        return style if style.name == name
      end
      nil
    end

    # Check if style exists in current document
    # @param name [String] name of style
    # @return [True, False]
    def style_exist?(name)
      !document_style_by_name(name).nil?
    end

    def self.parse(path)
      OOXMLDocumentObject.root_subfolder = 'word/'
      OOXMLDocumentObject.path_to_folder = path
      OOXMLDocumentObject.xmls_stack = []
      OOXMLDocumentObject.namespace_prefix = 'w'
      @comments = []
      @default_paragraph = DocxParagraph.new
      @default_character = DocxParagraphRun.new
      @default_table_properties = TableProperties.new
      PresentationTheme.parse('word/theme/theme1.xml')
      OOXMLDocumentObject.add_to_xmls_stack('word/styles.xml')
      doc = Nokogiri::XML(File.open(OOXMLDocumentObject.current_xml))
      doc.search('//w:docDefaults').each do |doc_defaults|
        doc_defaults.xpath('w:pPrDefault').each do |p_pr_defaults|
          @default_paragraph = DocxParagraph.parse(p_pr_defaults, 0)
        end
        doc_defaults.xpath('w:rPrDefault').each do |r_pr_defaults|
          r_pr_defaults.xpath('w:rPr').each do |r_pr|
            @default_character = DocxParagraphRun.parse(r_pr, DocxParagraphRun.new)
          end
        end
      end
      parse_default_style
      doc_structure = DocumentStructure.new
      doc_structure.document_styles = DocumentStyle.parse_list
      number = 0
      OOXMLDocumentObject.add_to_xmls_stack('word/document.xml')
      doc = Nokogiri::XML(File.open(OOXMLDocumentObject.current_xml))
      doc.search('//w:document').each do |document|
        document.xpath('w:background').each do |background|
          doc_structure.background = DocumentBackground.new
          doc_structure.background.color1 = Color.from_int16(background.attribute('color').value)
          background.xpath('v:background').each do |second_color|
            unless second_color.attribute('targetscreensize').nil?
              doc_structure.background.size = second_color.attribute('targetscreensize').value.sub(',', 'x')
            end
            second_color.xpath('v:fill').each do |fill|
              if !fill.attribute('color2').nil?
                doc_structure.background.color2 = Color.from_int16(fill.attribute('color2').value.split(' ').first.delete('#'))
                doc_structure.background.type = fill.attribute('type').value
              elsif !fill.attribute('id').nil?
                path_to_image = path + 'word/' + ParserDocxHelper.get_link_from_rels(fill.attribute('id').value, path)
                image_name = File.basename(path_to_image)
                doc_structure.background.image = path + File.basename(path) + '/' + image_name
                unless File.exist?(doc_structure.background.image.sub(File.basename(doc_structure.background.image), ''))
                  FileUtils.mkdir(doc_structure.background.image.sub(File.basename(doc_structure.background.image), ''))
                end
                FileUtils.cp path_to_image, doc_structure.background.image
                doc_structure.background.type = 'image'
              end
            end
          end
        end
        document.xpath('w:body').each do |body|
          body.xpath('*').each do |element|
            if element.name == 'p'
              child = element.child
              unless child.nil? && doc_structure.elements.last.class == Table
                paragraph_style = DocxParagraph.parse(element, number, @default_paragraph, @default_character)
                number += 1
                doc_structure.elements << paragraph_style.copy
              end
            elsif element.name == 'tbl'
              table = Table.parse(element, number, TableProperties.new, DocumentStructure.default_table_paragraph_style, DocumentStructure.default_table_run_style)
              number += 1
              doc_structure.elements << table
            end
          end
          body.xpath('w:sectPr').each do |sect_pr|
            doc_structure.page_properties = PageProperties.parse(sect_pr, @default_paragraph, @default_character)
            doc_structure.notes = doc_structure.page_properties.notes # keep copy of notes to compatibility with previous docx models
          end
        end
      end
      OOXMLDocumentObject.xmls_stack.pop
      doc_structure.document_properties = DocumentProperties.parse
      doc_structure.comments = Comment.parse_list
      doc_structure
    end

    def self.parse_default_style
      doc = Nokogiri::XML(File.open(OOXMLDocumentObject.path_to_folder + 'word/styles.xml'))
      doc.search('//w:style').each do |style|
        next if style.attribute('default').nil?
        if (style.attribute('default').value == '1' || style.attribute('default').value == 'on' || style.attribute('default').value == 'true') && style.attribute('type').value == 'paragraph'
          style.xpath('w:pPr').each do |paragraph_pr_tag|
            DocxParagraph.parse_paragraph_style(paragraph_pr_tag, @default_paragraph, @default_character)
          end
          style.xpath('w:rPr').each do |character_pr_tag|
            DocxParagraphRun.parse(character_pr_tag, @default_character, @default_character)
          end
        elsif (style.attribute('default').value == '1' || style.attribute('default').value == 'on' || style.attribute('default').value == 'true') && style.attribute('type').value == 'character'
          style.xpath('w:rPr').each do |character_pr_tag|
            DocxParagraphRun.parse(character_pr_tag, @default_character, @default_character)
          end
        end
      end
      DocumentStructure.default_table_paragraph_style = @default_paragraph.copy
      DocumentStructure.default_table_paragraph_style.spacing = Spacing.new(0, 0, 1, :auto)
      DocumentStructure.default_table_run_style = @default_character.copy
      doc.search('//w:style').each do |style|
        next if style.attribute('default').nil?
        next unless (style.attribute('default').value == '1' || style.attribute('default').value == 'on' || style.attribute('default').value == 'true') && style.attribute('type').value == 'table'
        style.xpath('w:tblPr').each do |table_pr_tag|
          @default_table_properties = TableProperties.parse(table_pr_tag)
        end
        style.xpath('w:pPr').each do |table_paragraph_pr_tag|
          DocxParagraph.parse_paragraph_style(table_paragraph_pr_tag, @default_table_paragraph, @default_table_character)
        end
        style.xpath('w:rPr').each do |table_character_pr_tag|
          DocxParagraphRun.parse(table_character_pr_tag, @default_table_character, @default_character)
        end
      end
    end

    class << self
      attr_accessor :default_table_run_style
      attr_accessor :default_table_paragraph_style
      attr_accessor :default_paragraph_style
    end
  end
end
