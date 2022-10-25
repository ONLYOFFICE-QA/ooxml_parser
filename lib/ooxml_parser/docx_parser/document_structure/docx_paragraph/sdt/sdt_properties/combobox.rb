# frozen_string_literal: true

require_relative 'combobox/list_item'
module OoxmlParser
  # Class for parsing `comboBox` tag
  class ComboBox < OOXMLDocumentObject
    # @return [Array<ListItem>] combobox items
    attr_reader :list_items

    def initialize(parent: nil)
      @list_items = []
      super
    end

    # Parse ComboBox object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ComboBox] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'listItem'
          @list_items << ListItem.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
