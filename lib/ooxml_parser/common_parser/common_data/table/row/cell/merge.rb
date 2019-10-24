# frozen_string_literal: true

module OoxmlParser
  class GridSpan
    attr_accessor :type, :count_of_merged_cells, :value

    def initialize(type = 'horizontal', value = nil, count_of_merged_cells = 2)
      @type = type
      @count_of_merged_cells = count_of_merged_cells
      @value = value
    end

    alias count_rows_in_span count_of_merged_cells

    # Parse Grid Span data
    # @param [Nokogiri::XML:Element] node with GridSpan data
    # @return [GridSpan] value of GridSpan
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'count_rows_in_span'
          @count_of_merged_cells = value.value
        when 'val'
          @value = value.value.to_i
        when 'type'
          @type = value.value.to_sym
        end
      end
      self
    end
  end
end
