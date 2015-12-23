# Document Properties
module OoxmlParser
  class DocumentProperties < OOXMLDocumentObject
    attr_accessor :pages, :words

    def initialize(pages = nil, words = nil)
      @pages = pages
      @words = words
    end

    # Parse Document properties
    # @return [DocumentProperties]
    def self.parse
      properties = DocumentProperties.new
      doc_props = XmlSimple.xml_in(File.open("#{OOXMLDocumentObject.path_to_folder}docProps/app.xml"))
      properties.pages = doc_props['Pages'].first.to_i unless doc_props['Pages'].nil?
      properties.words = doc_props['Words'].first.to_i unless doc_props['Words'].nil?
      properties
    end
  end
end
