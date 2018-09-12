module OoxmlParser
  # method to help to work with Presentation
  module PresentationHelpers
    # @return [True, false] if structure contain any user data
    def with_data?
      return true if @slides.length > 1

      @slides.each do |current_slide|
        return true if current_slide.with_data?
      end
      false
    end
  end
end
