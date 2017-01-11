module OoxmlParser
  # Class for storing Numbering Format data, `numFmt` tag
  class NumberingFormat < OOXMLDocumentObject
    # @return [String] value of start
    attr_accessor :value

    # Parse NumberingFormat
    # @param [Nokogiri::XML:Node] node with NumberingFormat
    # @return [NumberingFormat] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = value.value.to_s
        end
      end
      self
    end
  end
end
