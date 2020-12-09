# frozen_string_literal: true

require_relative 'cache_fields/cache_field'

module OoxmlParser
  # Class for parsing <cacheFields> tag
  class CacheFields < OOXMLDocumentObject
    # @return [Integer] count
    attr_reader :count
    # @return [Array<CacheField>] list of CacheField object
    attr_reader :cache_field

    def initialize(parent: nil)
      @cache_field = []
      super
    end

    # @return [CacheField] accessor
    def [](key)
      @cache_field[key]
    end

    # Parse `<cacheFields>` tag
    # @param [Nokogiri::XML:Element] node with cacheFields data
    # @return [cacheFields]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'count'
          @count = value.value.to_i
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'cacheField'
          @cache_field << CacheField.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
