# frozen_string_literal: true

require_relative 'tabs/tab'
module OoxmlParser
  # Class for working with `w:tabs`
  class Tabs < OOXMLDocumentObject
    # @return [Array, ParagraphTabs]
    attr_accessor :tabs_array

    def initialize(parent: nil)
      @tabs_array = []
      super
    end

    # @return [Array, Column] accessor
    def [](key)
      @tabs_array[key]
    end

    # @return [True, False] if empty
    def empty?
      @tabs_array.empty?
    end

    # Parse Tabs data
    # @param [Nokogiri::XML:Element] node with Tabs data
    # @return [Tabs] value of Columns data
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'tab'
          @tabs_array << Tab.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
