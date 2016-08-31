module OoxmlParser
  # Class for parsing `w:tblStyleColBandSize` object
  # When a style specifies the format for a band of columns in a
  # table (a set of contiguous columns),
  # this specifies the number of columns in a band.
  class TableStyleColumnBandSize < OOXMLDocumentObject
    # @return [Integer] value of table style column band size
    attr_accessor :value

    # Parse TableStyleColumnBandSize
    # @param [Nokogiri::XML:Node] node with TableStyleColumnBandSize
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
