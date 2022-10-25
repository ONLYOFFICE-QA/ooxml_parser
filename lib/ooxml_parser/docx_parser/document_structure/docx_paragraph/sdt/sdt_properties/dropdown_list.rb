# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `dropDownList` tag
  class DropdownList < OOXMLDocumentObject
    # @return [Array<ListItem>] dropdown items
    attr_reader :list_items

    def initialize(parent: nil)
      @list_items = []
      super
    end

    # Parse DropdownList object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DropdownList] result of parsing
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
