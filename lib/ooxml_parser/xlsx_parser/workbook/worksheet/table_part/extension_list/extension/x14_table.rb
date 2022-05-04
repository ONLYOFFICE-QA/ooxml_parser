# frozen_string_literal: true

module OoxmlParser
  # Class for `x14:table` data
  class X14Table < OOXMLDocumentObject
    # @return [String] alternate text for table
    attr_accessor :alt_text
    # @return [String] alternate text summary table
    attr_accessor :alt_text_summary

    # Parse X14Table data
    # @param [Nokogiri::XML:Element] node with X14Table data
    # @return [X14Table] value of X14Table data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'altText'
          @alt_text = value.value.to_s
        when 'altTextSummary'
          @alt_text_summary = value.value.to_s
        end
      end
      self
    end
  end
end
