module OoxmlParser
  class HeaderFooter < OOXMLDocumentObject
    attr_accessor :id, :description, :type

    def initialize(type = :header)
      @type = type
    end

    def parse(type, id, default_paragraph, default_character)
      if type == :header
        doc = Nokogiri::XML(File.open(OOXMLDocumentObject.path_to_folder + 'word/footnotes.xml'))
        doc.search('//w:footnote').each do |footnote|
          next unless footnote.attribute('id').value == id
          description = []
          paragraph_number = 0
          footnote.xpath('w:p').each do |paragraph|
            description << DocxParagraph.parse(paragraph, paragraph_number, default_paragraph, default_character)
            paragraph_number += 1
          end
          return description
        end
      elsif type == :header
        doc = Nokogiri::XML(File.open(OOXMLDocumentObject.path_to_folder + 'word/endnotes.xml'))
        doc.search('//w:endnote').each do |footnote|
          next unless footnote.attribute('id').value == id
          description = []
          paragraph_number = 0
          footnote.xpath('w:p').each do |paragraph|
            description << DocxParagraph.parse(paragraph, paragraph_number, default_paragraph, default_character)
            paragraph_number += 1
          end
          return description
        end
      end
    end
  end
end
