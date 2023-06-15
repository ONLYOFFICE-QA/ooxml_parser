# frozen_string_literal: true

module OoxmlParser
  # Class for `x14:dataField` data
  class X14DataField < OOXMLDocumentObject
    # @return [Symbol] pivot show as type
    attr_reader :pivot_show_as

    # Parse X14DataField data
    # @param [Nokogiri::XML:Element] node with X14DataField data
    # @return [X14Table] value of X14DataField data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'pivotShowAs'
          @pivot_show_as = value.value.to_sym
        end
      end
      self
    end
  end
end
