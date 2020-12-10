# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <pivotTableStyleInfo> tag
  class PivotTableStyleInfo < OOXMLDocumentObject
    # @return [String] name of style
    attr_reader :name
    # @return [True, False] show row header
    attr_reader :show_row_header
    # @return [True, False] show column header
    attr_reader :show_column_header
    # @return [True, False] show row stripes
    attr_reader :show_row_stripes
    # @return [True, False] show column stripes
    attr_reader :show_column_stripes
    # @return [True, False] show last column
    attr_reader :show_last_column

    # Parse `<pivotTableStyleInfo>` tag
    # @param [Nokogiri::XML:Element] node with PivotTableStyleInfo data
    # @return [PivotTableStyleInfo]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'name'
          @name = value.value.to_s
        when 'showRowHeaders'
          @show_row_header = attribute_enabled?(value)
        when 'showColHeaders'
          @show_column_header = attribute_enabled?(value)
        when 'showRowStripes'
          @show_row_stripes = attribute_enabled?(value)
        when 'showColStripes'
          @show_column_stripes = attribute_enabled?(value)
        when 'showLastColumn'
          @show_last_column = attribute_enabled?(value)
        end
      end
      self
    end
  end
end
