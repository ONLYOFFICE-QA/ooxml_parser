# frozen_string_literal: true

module OoxmlParser
  # Parsing common chart data
  class CommonChartData < OOXMLDocumentObject
    # @return [ValuedChild] type of grouping
    attr_reader :grouping
    # @return [Array<Series>] series data
    attr_reader :series

    def initialize(parent: nil)
      @series = []
      @parent = parent
    end

    # Parse CommonChartData object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [CommonChartData] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'grouping'
          @grouping = ValuedChild.new(:symbol, parent: self).parse(node_child)
        when 'ser'
          @series << Series.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
