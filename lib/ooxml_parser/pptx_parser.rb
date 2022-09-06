# frozen_string_literal: true

require_relative 'pptx_parser/presentation'

module OoxmlParser
  # Basic class for parsing pptx
  class PptxParser
    # Parse pptx file
    # @param path_to_file [String] file path
    # @return [Presentation] result of parse
    def self.parse_pptx(path_to_file)
      file = OoxmlFile.new(path_to_file)
      Parser.parse_format(file) do |yielded_file|
        Presentation.new(unpacked_folder: yielded_file.path_to_folder).parse
      end
    end
  end
end
