# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <sharedItems> tag
  class SharedItems < OOXMLDocumentObject
    # Parse `<sharedItems>` tag
    # # @param [Nokogiri::XML:Element] node with WorksheetSource data
    # @return [sharedItems]
    def parse(node)
      node.attributes.each do |key, value|
      end
      self
    end
  end
end
