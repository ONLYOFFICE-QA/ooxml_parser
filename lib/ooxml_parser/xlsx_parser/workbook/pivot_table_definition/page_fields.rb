# frozen_string_literal: true

require_relative 'page_fields/page_field'

module OoxmlParser
  # Class for parsing <pageFields> tag
  class PageFields < OOXMLDocumentObject
    # @return [Integer] count
    attr_reader :count
    # @return [Array<PageField>] list of PageField object
    attr_reader :page_field

    def initialize(parent: nil)
      @page_field = []
      super
    end

    # @return [PageField] accessor
    def [](key)
      @page_field[key]
    end

    # Parse `<pageFields>` tag
    # @param [Nokogiri::XML:Element] node with pageFields data
    # @return [PageFields]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'count'
          @count = value.value.to_i
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'pageField'
          @page_field << PageField.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
