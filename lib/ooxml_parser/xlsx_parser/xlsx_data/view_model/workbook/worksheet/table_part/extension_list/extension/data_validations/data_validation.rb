# frozen_string_literal: true

module OoxmlParser
  # Class for `dataValidation` data
  class DataValidation < OOXMLDocumentObject
    # @return [Boolean] should blank entries be valid
    attr_reader :allow_blank
    # @return [String] specifies the message text of the error alert
    attr_reader :error
    # @return [Symbol] type of error
    attr_reader :error_style
    # @return [String] the text of the title bar of the error alert
    attr_reader :error_title
    # @return [Symbol] Input Method Editor (IME) mode
    attr_reader :ime_mode
    # @return [Symbol] Relational operator used with this data validation
    attr_reader :operator
    # @return [Symbol] Specifies whether to display the drop-down combo box
    attr_reader :show_dropdown
    # @return [Symbol] Specifies whether to display the input prompt
    attr_reader :show_input_message
    # @return [Symbol] Specifies whether to display error alert message
    attr_reader :show_error_message
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
        when 'error'
          @error = value.value.to_s
        when 'errorStyle'
          @error_style = value.value.to_sym
        when 'errorTitle'
          @error_title = value.value.to_s
        when 'imeMode'
          @ime_mode = value.value.to_sym
        when 'operator'
          @operator = value.value.to_sym
        when 'type'
          @type = value.value.to_sym
        when 'showDropDown'
          @show_dropdown = attribute_enabled?(value)
        when 'showInputMessage'
          @show_input_message = attribute_enabled?(value)
        when 'showErrorMessage'
          @show_error_message = attribute_enabled?(value)
        when 'uid'
          @uid = value.value.to_s
        end
      end
      self
    end
  end
end