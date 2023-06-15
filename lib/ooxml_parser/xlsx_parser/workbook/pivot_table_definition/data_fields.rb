# frozen_string_literal: true

require_relative 'data_fields/data_field'

module OoxmlParser
  # Class for parsing <dataFields> tag
  class DataFields < OOXMLDocumentObject
    # @return [Integer] count
    attr_reader :count
    # @return [Array<DataField>] list of DataField object
    attr_reader :data_field

    def initialize(parent: nil)
      @data_field = []
      super
    end

    # @return [DataField] accessor
    def [](key)
      @data_field[key]
    end

    # Parse `<dataFields>` tag
    # @param [Nokogiri::XML:Element] node with dataFields data
    # @return [DataFields]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'count'
          @count = value.value.to_i
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'dataField'
          @data_field << DataField.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
