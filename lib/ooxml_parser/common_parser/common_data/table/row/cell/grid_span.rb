# frozen_string_literal: true

module OoxmlParser
  # Class for describing `gridSpan` tag
  class GridSpan < OOXMLDocumentObject
    # @return [String] type of grid span
    attr_reader :type
    # @return [String] value of grid span merged cells
    attr_reader :value

    def initialize(parent: nil)
      @type = 'horizontal'
      @parent = parent
    end

    alias count_of_merged_cells value

    extend Gem::Deprecate
    deprecate :type, 'GridSpan always horizontal', 2069, 1
    deprecate :count_of_merged_cells, 'value', 2069, 1

    # Parse Grid Span data
    # @param [Nokogiri::XML:Element] node with GridSpan data
    # @return [GridSpan] value of GridSpan
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = value.value.to_i
        end
      end
      self
    end
  end
end
