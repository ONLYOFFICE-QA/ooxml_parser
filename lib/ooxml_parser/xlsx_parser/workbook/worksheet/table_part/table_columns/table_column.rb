# frozen_string_literal: true

module OoxmlParser
  # Class for `tableColumn` tag data
  class TableColumn < OOXMLDocumentObject
    # @return [Integer] id of table column
    attr_reader :id
    # @return [String] name
    attr_reader :name
    # @return [String] total row label
    attr_reader :totals_row_label
    # @return [String] total row function
    attr_reader :totals_row_function

    # Parse Table Column data
    # @param [Nokogiri::XML:Element] node with TableColumn data
    # @return [TableColumn] value of TableColumn data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'id'
          @id = value.value.to_i
        when 'name'
          @name = value.value.to_s
        when 'totalsRowLabel'
          @totals_row_label = value.value.to_s
        when 'totalsRowFunction'
          @totals_row_function = value.value.to_s
        end
      end
      self
    end
  end
end
