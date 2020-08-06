# frozen_string_literal: true

require_relative 'xlsx_borders/xlsx_border'
module OoxmlParser
  # Class for parsing `borders`
  class XlsxBorders < OOXMLDocumentObject
    # @return [Integer] count of xlsx borders
    attr_reader :count
    # @return [Array<XlsxBorder>] list of xlsx borders
    attr_reader :borders_array

    def initialize(parent: nil)
      @borders_array = []
      super
    end

    # Parse XlsxBorders data
    # @param [Nokogiri::XML:Element] node with XlsxBorders data
    # @return [XlsxBorders] value of XlsxBorders data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'count'
          @count = value.value.to_i
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'border'
          @borders_array << XlsxBorder.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
