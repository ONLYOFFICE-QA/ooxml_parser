require_relative 'series/series_text'
module OoxmlParser
  # Class for parsing `c:ser` object
  class Series < OOXMLDocumentObject
    # @return [Index] index of chart
    attr_reader :index
    # @return [ValuedChild] order of chart
    attr_reader :order
    # @return [SeriesText] text of series
    attr_accessor :text
    # @return [Categories] categories of chart
    attr_accessor :categories
    # @return [DisplayLabelsProperties]
    attr_accessor :display_labels

    # Parse Series
    # @param [Nokogiri::XML:Node] node with Series
    # @return [Series] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'idx'
          @index = ValuedChild.new(:integer, parent: self).parse(node_child)
        when 'order'
          @order = ValuedChild.new(:integer, parent: self).parse(node_child)
        when 'tx'
          @text = SeriesText.new(parent: self).parse(node_child)
        when 'cat'
          @categories = Categories.new(parent: self).parse(node_child)
        when 'dLbls'
          @display_labels = DisplayLabelsProperties.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
