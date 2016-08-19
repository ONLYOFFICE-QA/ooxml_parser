module OoxmlParser
  # method to help to work with DocumentStructure
  module WorkbookHelpers
    # @return [True, false] if structure contain any user data
    def with_data?
      return true if @worksheets.length > 1
      return true if @worksheets.first.with_data?
      false
    end
  end
end
