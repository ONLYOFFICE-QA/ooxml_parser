# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <cacheSource> file
  class CacheSource < OOXMLDocumentObject
    # @return [String] type
    attr_reader :type

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
      self
    end
  end
end
