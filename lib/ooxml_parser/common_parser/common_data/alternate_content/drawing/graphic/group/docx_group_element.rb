module OoxmlParser
  # Docx Group Element data
  class DocxGroupElement < OOXMLDocumentObject
    attr_accessor :type, :object

    def initialize(type = nil, parent: nil)
      @type = type
      @parent = parent
    end
  end
end
