module OoxmlParser
  class DocumentGrid
    attr_accessor :type, :line_pitch, :char_space

    def initialize(type = nil, line_pitch = nil, char_space = nil)
      @type = type
      @line_pitch = line_pitch
      @char_space = char_space
    end
  end
end
