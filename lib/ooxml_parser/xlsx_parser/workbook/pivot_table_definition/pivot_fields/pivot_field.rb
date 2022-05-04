# frozen_string_literal: true

require_relative 'pivot_field/items'

module OoxmlParser
  # Class for parsing <pivotField> tag
  class PivotField < OOXMLDocumentObject
    # @return [String] axis value
    attr_reader :axis
    # @return [True, False] should show all
    attr_reader :show_all
    # @return [Items] contain item
    attr_reader :items

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

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'items'
          @items = Items.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
