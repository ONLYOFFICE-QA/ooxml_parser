module OoxmlParser
  # Parsing `numFmt` tag
  class NumberFormat < OOXMLDocumentObject
    # @return [Integer] id of number format
    attr_accessor :id
    # @return [String] code of format
    attr_accessor :format_code

    # Parse NumberFormat data
    # @param [Nokogiri::XML:Element] node with NumberFormat data
    # @return [NumberFormat] value of NumberFormat data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'numFmtId'
          @id = value.value.to_i
        when 'formatCode'
          @format_code = value.value.to_s
        end
      end
      self
    end
  end
end
