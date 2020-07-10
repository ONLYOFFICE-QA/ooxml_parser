# frozen_string_literal: true

require_relative 'pptx_data/presentation.rb'

module OoxmlParser
  # Basic class for parsing pptx
  class PptxParser
    # Parse pptx file
    # @param path_to_file [String] file path
    # @return [Presentation] result of parse
    def self.parse_pptx(path_to_file)
      Parser.parse_format(path_to_file) do
        Presentation.new.parse
      end
    end
  end
end
