# frozen_string_literal: true

module OoxmlParser
  # Class for `dataValidation` data
  class DataValidation < OOXMLDocumentObject
    # @return [String] uid of validation
    attr_reader :uid

    # Parse DataValidation data
    # @param [Nokogiri::XML:Element] node with DataValidation data
    # @return [DataValidation] value of DataValidation data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'uid'
          @uid = value.value.to_s
        end
      end
      self
    end
  end
end
