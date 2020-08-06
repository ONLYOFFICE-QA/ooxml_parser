# frozen_string_literal: true

require_relative 'filter_column/custom_filters'
module OoxmlParser
  # Class for `filterColumn` data
  # The filterColumn collection identifies a particular
  # column in the AutoFilter range and specifies
  # filter information that has been applied to this column.
  class FilterColumn < OOXMLDocumentObject
    # @return [True, False] Flag indicating whether the filter button is visible.
    attr_accessor :show_button
    # @return [CustomFilters] list of filters
    attr_reader :custom_filters

    def initialize(parent: nil)
      @show_button = true
      super
    end

    # Parse FilterColumn data
    # @param [Nokogiri::XML:Element] node with FilterColumn data
    # @return [FilterColumn] value of FilterColumn data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'showButton'
          @show_button = attribute_enabled?(value)
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'customFilters'
          @custom_filters = CustomFilters.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
