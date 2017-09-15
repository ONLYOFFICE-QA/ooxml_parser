module OoxmlParser
  # Class for `sparklineGroup` data
  class SparklineGroup < OOXMLDocumentObject
    # @return [Symbol] type of group
    attr_reader :type
    # @return [OoxmlSize] line weight
    attr_reader :line_weight
    # @return [True, False] show high point
    attr_reader :high_point
    # @return [True, False] show low point
    attr_reader :low_point
    # @return [True, False] show first point
    attr_reader :first_point
    # @return [True, False] show last point
    attr_reader :last_point
    # @return [True, False] show negative point
    attr_reader :negative_point
    # @return [True, False] show markers
    attr_reader :markers

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
        when 'high'
          @high_point = attribute_enabled?(value)
        when 'low'
          @low_point = attribute_enabled?(value)
        when 'first'
          @first_point = attribute_enabled?(value)
        when 'last'
          @last_point = attribute_enabled?(value)
        when 'negative'
          @negative_point = attribute_enabled?(value)
        when 'markers'
          @markers = attribute_enabled?(value)
        end
      end
      self
    end
  end
end
