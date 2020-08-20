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
      super
    end

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
