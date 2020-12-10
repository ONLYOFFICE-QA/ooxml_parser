# frozen_string_literal: true

require_relative 'pivot_fields/pivot_field'

module OoxmlParser
  # Class for parsing <pivotFields> tag
  class PivotFields < OOXMLDocumentObject
    # @return [Integer] count
    attr_reader :count
    # @return [Array<PivotField>] list of PivotField object
    attr_reader :pivot_field

    def initialize(parent: nil)
      @pivot_field = []
      super
    end

    # @return [PivotField] accessor
    def [](key)
      @pivot_field[key]
    end

    # Parse `<pivotFields>` tag
    # @param [Nokogiri::XML:Element] node with pivotFields data
    # @return [PivotFields]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'count'
          @count = value.value.to_i
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'pivotField'
          @pivot_field << PivotField.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
