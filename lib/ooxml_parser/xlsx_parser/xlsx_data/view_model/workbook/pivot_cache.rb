# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <pivotCache> tag
  class PivotCache < OOXMLDocumentObject
    # @return [Integer] cacheId of pivot cache
    attr_reader :cache_id
    # @return [String] id of pivot cache
    attr_reader :id

    def initialize(parent: nil)
      @parent = parent
    end

    # Parse Pivot Cache data
    # @param [Nokogiri::XML:Element] node with Pivot Cache data
    # @return [PivotCache] value of PivotCache
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'cacheId'
          @cache_id = value.value.to_i
        when 'id'
          @id = value.value.to_s
        end
      end
      self
    end
  end
end
