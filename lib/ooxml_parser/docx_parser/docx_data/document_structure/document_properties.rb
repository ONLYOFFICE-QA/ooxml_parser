module OoxmlParser
  # Document Properties
  class DocumentProperties < OOXMLDocumentObject
    attr_accessor :pages, :words

    # Parse Document properties
    # @return [DocumentProperties]
    def parse
      properties_file = "#{OOXMLDocumentObject.path_to_folder}docProps/app.xml"
      unless File.exist?(properties_file)
        warn "There is no 'docProps/app.xml' in docx. It may be some problem with it"
        return self
      end
      doc_props = XmlSimple.xml_in(File.open(properties_file))
      @pages = doc_props['Pages'].first.to_i unless doc_props['Pages'].nil?
      @words = doc_props['Words'].first.to_i unless doc_props['Words'].nil?
      self
    end
  end
end
