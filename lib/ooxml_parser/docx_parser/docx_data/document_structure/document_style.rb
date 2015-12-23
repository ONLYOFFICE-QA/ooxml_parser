module OoxmlParser
  # Class for describing styles containing in +styles.xml+
  class DocumentStyle < OOXMLDocumentObject
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
    # @return [True, False] Latent Style Primary Style Setting
    # Used to determine if current style is visible in style list in editors
    # According to http://www.wordarticles.com/Articles/WordStyles/LatentStyles.php
    attr_accessor :q_format
    alias_method :visible?, :q_format

    def initialize
      @q_format = false
    end

    # Parse single document style
    # @return [DocumentStyle]
    def self.parse(node)
      document_style = DocumentStyle.new
      node.attributes.each do |key, value|
        case key
        when 'type'
          document_style.type = value.value.to_sym
        when 'styleId'
          document_style.style_id = value.value.to_i
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
          document_style.run_properties = DocxParagraphRun.parse(subnode)
        when 'pPr'
          document_style.paragraph_properties = DocxParagraph.parse_paragraph_style(subnode)
        when 'qFormat'
          document_style.q_format = true
        end
      end
      document_style
    end

    # Parse all document style list
    # @return [Array, DocumentStyle]
    def self.parse_list
      styles_array = []
      doc = Nokogiri::XML(File.open(OOXMLDocumentObject.path_to_folder + 'word/styles.xml'))
      doc.search('//w:style').each do |style|
        styles_array << DocumentStyle.parse(style)
      end
      styles_array
    end
  end
end
