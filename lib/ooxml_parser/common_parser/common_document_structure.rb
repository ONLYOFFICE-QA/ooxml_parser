module OoxmlParser
  # Common document structure for DOCX, XLSX, PPTX file
  class CommonDocumentStructure < OOXMLDocumentObject
    # @return [String] path to original file
    attr_accessor :file_path
  end
end
