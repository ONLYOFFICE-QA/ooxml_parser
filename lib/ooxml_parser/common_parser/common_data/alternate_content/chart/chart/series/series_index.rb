module OoxmlParser
  # Class for parsing `c:idx` object
  class SeriesIndex < OOXMLDocumentObject
    # @return [Integer] value of index
    attr_accessor :value

    # Parse Index
    # @param [Nokogiri::XML:Node] node with Index
    # @return [Index] result of parsing
    def self.parse(node)
      index = SeriesIndex.new
      node.attributes.each do |key, value|
        case key
        when 'val'
          index.value = value.value.to_f
        end
      end
      index
    end
  end
end
