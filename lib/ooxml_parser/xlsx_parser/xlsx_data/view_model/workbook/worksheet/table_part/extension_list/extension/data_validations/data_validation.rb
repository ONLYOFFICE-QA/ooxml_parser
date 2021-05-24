# frozen_string_literal: true

module OoxmlParser
  # Class for `dataValidation` data
  class DataValidation < OOXMLDocumentObject
    # @return [Boolean] should blank entries be valid
    attr_reader :allow_blank
    # @return [Symbol] type of error
    attr_reader :error_style
    # @return [Symbol] type of validation
    attr_reader :type
    # @return [String] uid of validation
    attr_reader :uid

    # Parse DataValidation data
    # @param [Nokogiri::XML:Element] node with DataValidation data
    # @return [DataValidation] value of DataValidation data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'allowBlank'
          @allow_blank = attribute_enabled?(value)
        when 'errorStyle'
          @error_style = value.value.to_sym
        when 'type'
          @type = value.value.to_sym
        when 'uid'
          @uid = value.value.to_s
        end
      end
      self
    end
  end
end
