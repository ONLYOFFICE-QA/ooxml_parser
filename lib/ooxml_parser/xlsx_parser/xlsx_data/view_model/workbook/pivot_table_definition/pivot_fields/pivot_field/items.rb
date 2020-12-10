# frozen_string_literal: true

require_relative 'items/item'

module OoxmlParser
  # Class for parsing <items> tag
  class Items < OOXMLDocumentObject
    # @return [Integer] count
    attr_reader :count
    # @return [Array<Item>] list of items
    attr_reader :items

    def initialize(parent: nil)
      @items = []
      super
    end

    # @return [Item] accessor
    def [](key)
      @items[key]
    end

    # Parse `<items>` tag
    # @param [Nokogiri::XML:Element] node with items data
    # @return [Items]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'count'
          @count = value.value.to_i
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'item'
          @items << Item.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
