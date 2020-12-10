# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <pivotField> tag
  class PivotField < OOXMLDocumentObject
    # @return [String] axis value
    attr_reader :axis
    # @return [True, False] should show all
    attr_reader :show_all

    # Parse `<pivotField>` tag
    # # @param [Nokogiri::XML:Element] node with PivotField data
    # @return [PivotField]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'axis'
          @axis = value.value.to_s
        when 'showAll'
          @show_all = attribute_enabled?(value)
        end
      end
      self
    end
  end
end
