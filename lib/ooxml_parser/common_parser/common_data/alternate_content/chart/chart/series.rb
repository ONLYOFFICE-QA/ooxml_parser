require_relative 'series/series_index'
module OoxmlParser
  # Class for parsing `c:ser` object
  class Series < OOXMLDocumentObject
    # @return [Index] index of chart
    attr_accessor :index

    # Parse Series
    # @param [Nokogiri::XML:Node] node with Series
    # @return [Series] result of parsing
    def self.parse(node)
      series = Series.new
      node.xpath('*').each do |series_childe_node|
        case series_childe_node.name
        when 'idx'
          series.index = SeriesIndex.parse(series_childe_node)
        end
      end
      series
    end
  end
end
