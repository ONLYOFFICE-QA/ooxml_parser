# frozen_string_literal: true

require_relative 'custom_filters/custom_filter'
module OoxmlParser
  # Class for `CustomFilters` data
  class CustomFilters < OOXMLDocumentObject
    # @return [Integer] and value
    attr_reader :and
    # @return [Array<CustomFilter>] list of filters array
    attr_reader :filters_array

    def initialize(parent: nil)
      @filters_array = []
      @parent = parent
    end

    # @return [Array, CustomFilter] accessor
    def [](key)
      @filters_array[key]
    end

    # Parse CustomFilters data
    # @param [Nokogiri::XML:Element] node with CustomFilters data
    # @return [CustomFilters] value of CustomFilters data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'and'
          @and = value.value.to_i
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'customFilter'
          @filters_array << CustomFilter.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
