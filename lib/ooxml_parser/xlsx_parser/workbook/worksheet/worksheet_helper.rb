# frozen_string_literal: true

module OoxmlParser
  # Module with helper methods for worksheet
  module WorksheetHelper
    # @return [True, False] is columns are with default rules
    def default_columns?
      return true if @columns.empty?
      return true unless @columns.first.custom_width

      false
    end
  end
end
