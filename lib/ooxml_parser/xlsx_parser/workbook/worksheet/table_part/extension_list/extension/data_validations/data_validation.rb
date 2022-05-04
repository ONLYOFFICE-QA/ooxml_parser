# frozen_string_literal: true

require_relative 'data_validation/data_validation_formula'
module OoxmlParser
  # Class for `dataValidation` data
  class DataValidation < OOXMLDocumentObject
    # @return [Boolean] should blank entries be valid
    attr_reader :allow_blank
    # @return [String] Specifies the message text of the error alert
    attr_reader :error
    # @return [Symbol] Type of error
    attr_reader :error_style
    # @return [String] The text of the title bar of the error alert
    attr_reader :error_title
    # @return [DataValidationFormula] first formula of data validation
    attr_reader :formula1
    # @return [DataValidationFormula] second formula of data validation
    attr_reader :formula2
    # @return [Symbol] Input Method Editor (IME) mode
    attr_reader :ime_mode
    # @return [Symbol] Relational operator used with this data validation
    attr_reader :operator
    # @return [String] Message text of the input prompt
    attr_reader :prompt
    # @return [String] Text of the title bar of the input prompt
    attr_reader :prompt_title
    # @return [String] Ranges to which data validation is applied
    attr_reader :reference_sequence
    # @return [Symbol] Specifies whether to display the drop-down combo box
    attr_reader :show_dropdown
    # @return [Symbol] Specifies whether to display the input prompt
    attr_reader :show_input_message
    # @return [Symbol] Specifies whether to display error alert message
    attr_reader :show_error_message
    # @return [Symbol] Type of validation
    attr_reader :type
    # @return [String] UID of validation
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
        when 'prompt'
          @prompt = value.value.to_s
        when 'promptTitle'
          @prompt_title = value.value.to_s
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

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'formula1'
          @formula1 = DataValidationFormula.new(parent: self).parse(node_child)
        when 'formula2'
          @formula2 = DataValidationFormula.new(parent: self).parse(node_child)
        when 'sqref'
          @reference_sequence = node_child.text
        end
      end
      self
    end
  end
end
