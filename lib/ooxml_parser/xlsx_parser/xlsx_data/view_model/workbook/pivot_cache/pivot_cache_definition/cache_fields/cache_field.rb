# frozen_string_literal: true

require_relative 'cache_field/shared_items'

module OoxmlParser
  # Class for parsing <cacheField> tag
  class CacheField < OOXMLDocumentObject
    # @return [String] name of field
    attr_reader :name
    # @return [Integer] number format id
    attr_reader :number_format_id
    # @return [SharedItems] shared items
    attr_reader :shared_items

    # Parse `<cacheField>` tag
    # # @param [Nokogiri::XML:Element] node with WorksheetSource data
    # @return [CacheField]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'name'
          @name = value.value.to_s
        when 'numFmtId'
          @number_format_id = value.value.to_i
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'sharedItems'
          @shared_items = SharedItems.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
