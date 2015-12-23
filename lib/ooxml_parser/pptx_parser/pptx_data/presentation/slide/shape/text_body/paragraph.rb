module OoxmlParser
  class PresentationParagraph
    attr_accessor :properties, :characters, :text_field, :formulas

    def initialize(characters = [], formulas = [])
      @characters = characters
      @formulas = formulas
    end
  end
end
