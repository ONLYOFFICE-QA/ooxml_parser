# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <location> tag
  class Location < OOXMLDocumentObject
    # @return [String] ref of location
    attr_reader :ref
    # @return [True, False] first header row
    attr_reader :first_header_row
    # @return [True, False] first data row
    attr_reader :first_data_row
    # @return [True, False] first data column
    attr_reader :first_data_column

    # Parse `<location>` tag
    # @param [Nokogiri::XML:Element] node with location data
    # @return [Location]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'ref'
          @ref = value.value.to_s
        when 'firstHeaderRow'
          @first_header_row = attribute_enabled?(value)
        when 'firstDataRow'
          @first_data_row = attribute_enabled?(value)
        when 'firstDataCol'
          @first_data_column = attribute_enabled?(value)
        end
      end
      self
    end
  end
end
