# frozen_string_literal: true

require_relative 'column_row_items/column_row_item'

module OoxmlParser
  # Class for parsing <rowItems> and <colItems> tag
  class ColumnRowItems < OOXMLDocumentObject
    # @return [Integer] count of items
    attr_reader :count
    # @return [Array<RowColumnItem>] list of RowColumnItem object
    attr_reader :items

    def initialize(parent: nil)
      @items = []
      super
    end

    # @return [RowColumnItem] accessor
    def [](key)
      @items[key]
    end

    # Parse `<cacheField>` tag
    # # @param [Nokogiri::XML:Element] node with WorksheetSource data
    # @return [CacheField]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'count'
          @count = value.value.to_i
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'i'
          @items << ColumnRowItem.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
