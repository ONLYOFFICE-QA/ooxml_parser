# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <cacheField> tag
  class CacheField < OOXMLDocumentObject
    # Parse `<worksheetSource>` tag
    # # @param [Nokogiri::XML:Element] node with WorksheetSource data
    # @return [WorksheetSource]
    def parse(node)
      node.attributes.each do |key, value|
      end
      self
    end
  end
end
