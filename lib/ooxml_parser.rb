require_relative 'ooxml_parser/common_parser/common_data/ooxml_document_object'
require_relative 'ooxml_parser/common_parser/common_data/color'
require_relative 'ooxml_parser/common_parser/common_data/coordinates'
require_relative 'ooxml_parser/common_parser/common_data/font_style'
require_relative 'ooxml_parser/common_parser/common_data/borders_properties'
require_relative 'ooxml_parser/common_parser/common_data/hyperlink'
require_relative 'ooxml_parser/common_parser/common_data/paragraph'
require_relative 'ooxml_parser/common_parser/common_data/alternate_content/alternate_content'
require_relative 'ooxml_parser/common_parser/common_data/colors/presentation_fill'
require_relative 'ooxml_parser/common_parser/common_data/table'
require_relative 'ooxml_parser/common_parser/common_data/relationships'
require_relative 'ooxml_parser/common_parser/parser'
require_relative 'ooxml_parser/common_parser/common_document_structure'
require_relative 'ooxml_parser/configuration'
require_relative 'ooxml_parser/helpers/string_helper'
require_relative 'ooxml_parser/docx_parser/docx_parser'
require_relative 'ooxml_parser/xlsx_parser/xlsx_parser'
require_relative 'ooxml_parser/pptx_parser/pptx_parser'

module OoxmlParser
  class << self
    attr_writer :configuration
  end

  def self.configuration
    @configuration ||= Configuration.new
  end

  def self.reset
    @configuration = Configuration.new
  end

  def self.configure
    yield(configuration)
  end
end
