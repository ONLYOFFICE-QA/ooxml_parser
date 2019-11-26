# frozen_string_literal: true

module OoxmlParser
  # Fallback DOCX group element data
  class OldDocxGroupElement
    attr_accessor :type, :object

    def initialize(type = nil)
      @type = type
    end
  end
end
