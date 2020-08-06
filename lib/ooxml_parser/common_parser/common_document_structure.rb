# frozen_string_literal: true

require_relative 'common_data/content_types'
module OoxmlParser
  # Common document structure for DOCX, XLSX, PPTX file
  class CommonDocumentStructure < OOXMLDocumentObject
    # @return [String] path to original file
    attr_accessor :file_path
    # @return [Integer] default font size
    attr_reader :default_font_size
    # @return [Integer] default font typeface
    attr_reader :default_font_typeface
    # @return [FontStyle] Default font style of presentation
    attr_accessor :default_font_style
    # @return [ContentTypes] data about content types
    attr_accessor :content_types

    def initialize(params = {})
      @default_font_size = params.fetch(:default_font_size, 18)
      @default_font_typeface = params.fetch(:default_font_typeface, 'Arial')
      @default_font_style = FontStyle.new
      super(parent: nil)
    end
  end
end
