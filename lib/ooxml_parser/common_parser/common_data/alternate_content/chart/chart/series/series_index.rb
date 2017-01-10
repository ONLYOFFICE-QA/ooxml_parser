module OoxmlParser
  # Class for parsing `c:idx` object
  class SeriesIndex < OOXMLDocumentObject
    # @return [Integer] value of index
    attr_accessor :value

    # Parse SeriesIndex
    # @param [Nokogiri::XML:Node] node with Index
    # @return [SeriesIndex] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = value.value.to_f
        end
      end
      self
    end
  end
end
