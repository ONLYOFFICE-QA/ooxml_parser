module OoxmlParser
  # Class for `sparklineGroup` data
  class SparklineGroup < OOXMLDocumentObject
    # @return [Symbol] type of group
    attr_reader :type
    # @return [OoxmlSize] line weight
    attr_reader :line_weight

    # Parse SparklineGroup
    # @param [Nokogiri::XML:Node] node with SparklineGroup
    # @return [SparklineGroup] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'type'
          @type = value_to_symbol(value)
        when 'lineWeight'
          @line_weight = OoxmlSize.new(value.value.to_f, :point)
        end
      end
      self
    end
  end
end
