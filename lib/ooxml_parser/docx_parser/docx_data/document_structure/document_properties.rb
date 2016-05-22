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
      properties_file = "#{OOXMLDocumentObject.path_to_folder}docProps/app.xml"
      unless File.exist?(properties_file)
        warn "There is no 'docProps/app.xml' in docx. It may be some problem with it"
        return properties
      end
      doc_props = XmlSimple.xml_in(File.open(properties_file))
      properties.pages = doc_props['Pages'].first.to_i unless doc_props['Pages'].nil?
      properties.words = doc_props['Words'].first.to_i unless doc_props['Words'].nil?
      properties
    end
  end
end
