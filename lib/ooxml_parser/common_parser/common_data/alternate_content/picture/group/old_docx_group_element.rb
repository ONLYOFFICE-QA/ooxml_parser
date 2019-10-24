# frozen_string_literal: true

# Fallback DOCX Gropu Element data
module OoxmlParser
  class OldDocxGroupElement
    attr_accessor :type, :object

    def initialize(type = nil)
      @type = type
    end
  end
end
