# frozen_string_literal: true

require_relative 'row_fields/field'

module OoxmlParser
  # Class for parsing <rowFields> tag
  class RowFields < OOXMLDocumentObject
    # @return [Integer] count
    attr_reader :count
    # @return [Array<Field>] list of Field object
    attr_reader :fields

    def initialize(parent: nil)
      @fields = []
      super
    end

    # @return [Field] accessor
    def [](key)
      @fields[key]
    end

    # Parse `<rowFields>` tag
    # @param [Nokogiri::XML:Element] node with rowFields data
    # @return [RowFields]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'count'
          @count = value.value.to_i
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'field'
          @fields << Field.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
