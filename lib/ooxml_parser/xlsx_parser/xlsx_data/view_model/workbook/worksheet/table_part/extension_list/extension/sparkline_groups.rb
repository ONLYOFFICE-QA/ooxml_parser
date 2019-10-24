# frozen_string_literal: true

require_relative 'sparkline_groups/sparkline_group'
module OoxmlParser
  # Class for `sparklineGroups` data
  class SparklineGroups < OOXMLDocumentObject
    # @return [Array<SparklineGroup>] list of sparkline group
    attr_reader :sparklines_groups

    def initialize(parent: nil)
      @sparklines_groups = []
      @parent = parent
    end

    # @return [SparklineGroup] accessor
    def [](key)
      sparklines_groups[key]
    end

    # Parse SparklineGroups data
    # @param [Nokogiri::XML:Element] node with SparklineGroups data
    # @return [SparklineGroups] value of SparklineGroups data
    def parse(node)
      node.xpath('*').each do |column_node|
        case column_node.name
        when 'sparklineGroup'
          @sparklines_groups << SparklineGroup.new(parent: self).parse(column_node)
        end
      end
      self
    end
  end
end
