# frozen_string_literal: true

module OoxmlParser
  # method to help to work with DocumentStructure
  module DocumentStructureHelpers
    # @return [True, false] if structure contain any user data
    def with_data?
      @elements.each do |current_element|
        return true if current_element.with_data?
      end
      @notes.each do |current_note|
        current_note.elements.each do |note_element|
          return true if note_element.with_data?
        end
      end
      false
    end
  end
end
