require_relative 'pptx_data/presentation.rb'

module OoxmlParser
  class PptxParser
    def self.parse_pptx(path_to_file)
      Parser.parse(path_to_file) do
        Presentation.parse
      end
    end
  end
end
