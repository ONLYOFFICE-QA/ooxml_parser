module OoxmlParser
  # Class for parsing `w:trHeight` object
  class TableRowHeight < OOXMLDocumentObject
    # @return [Integer] value of size
    attr_accessor :value

    # Parse TableRowHeight
    # @param [Nokogiri::XML:Node] node with TableRowHeight
    # @return [TableRowHeight] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = OoxmlSize.new(value.value.to_f)
        end
      end
      self
    end
  end
end
