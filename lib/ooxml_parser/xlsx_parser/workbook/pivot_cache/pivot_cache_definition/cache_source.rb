# frozen_string_literal: true

require_relative 'cache_source/worksheet_source'

module OoxmlParser
  # Class for parsing <cacheSource> file
  class CacheSource < OOXMLDocumentObject
    # @return [String] type
    attr_reader :type
    # @return [WorksheetSource] source of worksheet data
    attr_reader :worksheet_source

    # Parse `<cacheSource>` tag
    # # @param [Nokogiri::XML:Element] node with CacheSource data
    # @return [CacheSource]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'type'
          @type = value.value.to_s
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'worksheetSource'
          @worksheet_source = WorksheetSource.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
