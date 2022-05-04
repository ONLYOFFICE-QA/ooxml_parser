# frozen_string_literal: true

require_relative 'data_validations/data_validation'

module OoxmlParser
  # Class for `dataValidations` data
  class DataValidations < OOXMLDocumentObject
    # @return [Integer] count of validations
    attr_reader :count
    # @return [Array<DataValidation>] list of data validations
    attr_reader :data_validations
    # @return [Boolean] is prompts disabled
    attr_reader :disable_prompts

    def initialize(parent: nil)
      @data_validations = []
      super
    end

    # @return [SparklineGroup] accessor
    def [](key)
      data_validations[key]
    end

    # Parse DataValidations data
    # @param [Nokogiri::XML:Element] node with DataValidations data
    # @return [DataValidations] value of DataValidations data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'count'
          @count = value.value.to_i
        when 'disablePrompts'
          @disable_prompts = attribute_enabled?(value)
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'dataValidation'
          @data_validations << DataValidation.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
