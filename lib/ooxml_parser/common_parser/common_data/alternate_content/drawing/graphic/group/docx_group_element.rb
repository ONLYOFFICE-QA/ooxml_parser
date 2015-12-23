# Docx Gropu Element data
module OoxmlParser
  class DocxGroupElement
    attr_accessor :type, :object

    def initialize(type = nil)
      @type = type
    end
  end
end
