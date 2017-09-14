module OoxmlParser
  # Class for `sparklineGroup` data
  class SparklineGroup < OOXMLDocumentObject
    # @return [Symbol] type of group
    attr_reader :type

    # Parse SparklineGroup
    # @param [Nokogiri::XML:Node] node with SparklineGroup
    # @return [SparklineGroup] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'type'
          @type = value_to_symbol(value)
        end
      end
      self
    end
  end
end
