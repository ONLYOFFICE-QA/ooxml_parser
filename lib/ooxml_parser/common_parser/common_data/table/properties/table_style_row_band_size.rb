# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `w:tblStyleRowBandSize` object
  # When a style specifies the format for a band for
  # rows in a table (a set of contiguous rows), this
  # specifies the number of rows in a band.
  class TableStyleRowBandSize < OOXMLDocumentObject
    # @return [Integer] value of table style column band size
    attr_accessor :value

    # Parse TableStyleRowBandSize
    # @param [Nokogiri::XML:Node] node with TableStyleRowBandSize
    # @return [TableStyleColumnBandSize] result of parsing
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
