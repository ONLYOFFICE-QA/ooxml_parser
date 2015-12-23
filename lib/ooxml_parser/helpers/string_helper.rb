# Class for working with strings
module OoxmlParser
  class StringHelper
    class << self
      # This method check if string contains numeric
      # @return [True, False] result of comparison
      def numeric?(string)
        true if Float(string)
      rescue ArgumentError
        false
      end

      # Check if string is complex number
      # @return [True, False] result of comparison
      def complex?(string)
        true if Complex(string.tr(',', '.'))
      rescue ArgumentError
        false
      end
    end
  end
end
