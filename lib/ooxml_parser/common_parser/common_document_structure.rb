module OoxmlParser
  # Common document structure for DOCX, XLSX, PPTX file
  class CommonDocumentStructure < OOXMLDocumentObject
    # @return [String] path to original file
    attr_accessor :file_path
    # @return [Integer] default font size
    attr_reader :default_font_size
    # @return [Integer] default font typeface
    attr_reader :default_font_typeface

    def initialize(params = {})
      @default_font_size = params.fetch(:default_font_size, 18)
      @default_font_typeface = params.fetch(:default_font_typeface, 'Arial')
    end
  end
end
