# frozen_string_literal: true

require_relative 'pivot_cache/pivot_cache_definition'

module OoxmlParser
  # Class for parsing <pivotCache> tag
  class PivotCache < OOXMLDocumentObject
    # @return [Integer] cacheId of pivot cache
    attr_reader :cache_id
    # @return [String] id of pivot cache
    attr_reader :id
    # @return [PivotCacheDefinition] parsed pivot cache definition
    attr_reader :pivot_cache_definition

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
      parse_pivot_cache_definition
      self
    end

    private

    # @return [PivotCacheDefinition] pivot cache definition for current pivot cache
    def parse_pivot_cache_definition
      definition_file = root_object.relationships.target_by_id(id)
      full_file_path = "#{OOXMLDocumentObject.path_to_folder}/xl/#{definition_file}"
      @pivot_cache_definition = PivotCacheDefinition.new(parent: root_object)
                                                    .parse(full_file_path)
    end
  end
end
