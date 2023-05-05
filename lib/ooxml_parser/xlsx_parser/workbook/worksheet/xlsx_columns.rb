# frozen_string_literal: true

require_relative 'xlsx_columns/xlsx_column_properties'
module OoxmlParser
  # Class for parsing XlsxColumns object <cols>
  class XlsxColumns < OOXMLDocumentObject
    # @return [Array<XlsxColumnProperties>] list of elements
    attr_reader :elements

    def initialize(parent: nil)
      @elements = []
      super
    end

    # Parse XlsxColumns
    # @param node [Nokogiri::XML::Element] node to parse
    # @return [XlsxColumns] value of TimeNodeList
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'col'
          @elements << XlsxColumnProperties.new(parent: parent).parse(node_child)
        end
      end
      self
    end
  end
end
