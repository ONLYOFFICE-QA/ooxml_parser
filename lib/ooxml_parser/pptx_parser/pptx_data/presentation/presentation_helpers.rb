module OoxmlParser
  # method to help to work with Presentation
  module PresentationHelpers
    # @return [True, false] if structure contain any user data
    def with_data?
      return true if @slides.length > 1
      @slides.first.elements.each do |current_element|
        return true if current_element.with_data?
      end
      false
    end
  end
end
