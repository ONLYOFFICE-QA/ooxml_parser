# frozen_string_literal: true

require_relative 'series/series_text'
require_relative 'series/x_y_values'
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
    # @return [XYValues] values of series
    attr_reader :values
    # @return [XYValues] values of x
    attr_reader :x_values
    # @return [XYValues] values of y
    attr_reader :y_values

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
        when 'val'
          @values = XYValues.new(parent: self).parse(node_child)
        when 'xVal'
          @x_values = XYValues.new(parent: self).parse(node_child)
        when 'yVal'
          @y_values = XYValues.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
