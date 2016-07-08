require_relative 'series/order'
require_relative 'series/series_index'
require_relative 'series/series_text'
module OoxmlParser
  # Class for parsing `c:ser` object
  class Series < OOXMLDocumentObject
    # @return [Index] index of chart
    attr_accessor :index
    # @return [Order] order of chart
    attr_accessor :order
    # @return [SeriesText] text of series
    attr_accessor :text

    # Parse Series
    # @param [Nokogiri::XML:Node] node with Series
    # @return [Series] result of parsing
    def self.parse(node)
      series = Series.new
      node.xpath('*').each do |series_childe_node|
        case series_childe_node.name
        when 'idx'
          series.index = SeriesIndex.parse(series_childe_node)
        when 'order'
          series.order = Order.parse(series_childe_node)
        when 'tx'
          series.text = SeriesText.parse(series_childe_node)
        end
      end
      series
    end
  end
end
