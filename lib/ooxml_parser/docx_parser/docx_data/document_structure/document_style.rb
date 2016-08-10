require_relative 'document_style/document_style_helper'
module OoxmlParser
  # Class for describing styles containing in +styles.xml+
  class DocumentStyle < OOXMLDocumentObject
    include TableStylePropertiesHelper
    # @return [Symbol] Type of style (+:paragraph+ or +:table+)
    attr_accessor :type
    # @return [FixNum] number of style
    attr_accessor :style_id
    # @return [String] name of style
    attr_accessor :name
    # @return [FixNum] id of style on which this style is based
    attr_accessor :based_on
    # @return [FixNum] id of next style
    attr_accessor :next_style
    # @return [DocxParagraphRun] run properties
    attr_accessor :run_properties
    # @return [DocxParagraph] run properties
    attr_accessor :paragraph_properties
    # @return [TableProperties] properties of table
    attr_accessor :table_properties
    # @return [Array, TableStyleProperties] list of table style properties
    attr_accessor :table_style_properties_list
    # @return [True, False] Latent Style Primary Style Setting
    # Used to determine if current style is visible in style list in editors
    # According to http://www.wordarticles.com/Articles/WordStyles/LatentStyles.php
    attr_accessor :q_format
    alias visible? q_format

    def initialize
      @q_format = false
      @table_style_properties_list = []
    end

    # Parse single document style
    # @param [OoxmlParser::OOXMLDocumentObject] parent parent object
    # @return [DocumentStyle]
    def self.parse(node, parent)
      document_style = DocumentStyle.new
      node.attributes.each do |key, value|
        case key
        when 'type'
          document_style.type = value.value.to_sym
        when 'styleId'
          document_style.style_id = value.value
        end
      end
      node.xpath('*').each do |subnode|
        case subnode.name
        when 'name'
          document_style.name = subnode.attribute('val').value
        when 'basedOn'
          document_style.based_on = subnode.attribute('val').value.to_i
        when 'next'
          document_style.next_style = subnode.attribute('val').value.to_i
        when 'rPr'
          document_style.run_properties = RunPropertiesDocument.parse(subnode)
        when 'pPr'
          document_style.paragraph_properties = DocxParagraph.parse_paragraph_style(subnode, parent: document_style)
        when 'tblPr'
          document_style.table_properties = TableProperties.parse(subnode, parent: document_style)
        when 'tblStylePr'
          document_style.table_style_properties_list << TableStyleProperties.parse(subnode, parent: document_style)
        when 'qFormat'
          document_style.q_format = true
        end
      end
      document_style.parent = parent
      document_style
    end

    # Parse all document style list
    # @return [Array, DocumentStyle]
    def self.parse_list(parent)
      styles_array = []
      doc = Nokogiri::XML(File.open(OOXMLDocumentObject.path_to_folder + 'word/styles.xml'))
      doc.search('//w:style').each do |style|
        styles_array << DocumentStyle.parse(style, parent)
      end
      styles_array
    end
  end
end
