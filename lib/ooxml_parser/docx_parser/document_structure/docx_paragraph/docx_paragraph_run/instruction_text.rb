# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `w:instrText` object
  class InstructionText < TextValue
    # @return [Boolean] is current object for hyperlink
    def hyperlink?
      value.include?('HYPERLINK')
    end

    # @return [Boolean] is current object for page number
    def page_number?
      value.match?(/PAGE\s+\\\*/)
    end

    # @return [Hyperlink] convert InstructionText to Hyperlink
    def to_hyperlink
      Hyperlink.new(value.sub('HYPERLINK ', '').split(' \\o ').first,
                    value.sub('HYPERLINK', '').split(' \\o ').last)
    end
  end
end
