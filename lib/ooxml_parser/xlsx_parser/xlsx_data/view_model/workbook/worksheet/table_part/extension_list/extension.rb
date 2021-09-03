# frozen_string_literal: true

require_relative 'extension/data_validations'
require_relative 'extension/sparkline_groups'
require_relative 'extension/x14_table'
require_relative 'extension/conditional_formatting'
module OoxmlParser
  # Class for `ext` data
  class Extension < OOXMLDocumentObject
    # @return [DataValidations] list of data validations
    attr_accessor :data_validations
    # @return [Array, ConditionalFormatting] array of conditional formattings
    attr_reader :conditional_formattings
    # @return [X14Table] table data in x14 namespace
    attr_accessor :table
    # @return [SparklineGroups] list of groups
    attr_reader :sparkline_groups

    def initialize(parent: nil)
      @conditional_formattings = []
      super
    end

    # Parse Extension data
    # @param [Nokogiri::XML:Element] node with Extension data
    # @return [Extension] value of Extension data
    def parse(node)
      node.xpath('*').each do |column_node|
        case column_node.name
        when 'dataValidations'
          @data_validations = DataValidations.new(parent: self).parse(column_node)
        when 'conditionalFormattings'
          column_node.xpath('*').each do |conditional_formatting_node|
            @conditional_formattings << ConditionalFormatting.new(parent: self).parse(conditional_formatting_node)
          end
        when 'table'
          @table = X14Table.new(parent: self).parse(column_node)
        when 'sparklineGroups'
          @sparkline_groups = SparklineGroups.new(parent: self).parse(column_node)
        end
      end
      self
    end
  end
end
