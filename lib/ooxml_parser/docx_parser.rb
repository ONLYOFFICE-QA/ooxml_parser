# frozen_string_literal: true

require_relative 'common_parser/common_data/valued_child'
require_relative 'docx_parser/document_structure'

module OoxmlParser
  # Basic class for DocxParser
  class DocxParser
    # Parse docx file
    # @param path_to_file [String] file path
    # @return [DocumentStructure] result of parse
    def self.parse_docx(path_to_file)
      Parser.parse_format(path_to_file) do |file|
        DocumentStructure.new(unpacked_folder: file.path_to_folder).parse
      end
    end
  end
end
