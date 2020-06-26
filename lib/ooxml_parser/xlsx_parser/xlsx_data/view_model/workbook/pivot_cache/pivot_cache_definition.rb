# frozen_string_literal: true

require_relative 'pivot_cache_definition/cache_source'
module OoxmlParser
  # Class for parsing <pivotCacheDefinition> file
  class PivotCacheDefinition < OOXMLDocumentObject
    # @return [String] id of pivot cache definition
    attr_reader :id
    # @return [CacheSource] source of pivot cache
    attr_reader :cache_source

    # Parse PivotCacheDefinition file
    # @param file [String] path to file
    # @return [PivotCacheDefinition]
    def parse(file)
      return nil unless File.exist?(file)

      document = parse_xml(file)
      node = document.xpath('*').first

      node.attributes.each do |key, value|
        case key
        when 'id'
          @id = value.value.to_s
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'cacheSource'
          @cache_source = CacheSource.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
