# frozen_string_literal: true

module OoxmlParser
  # Class for parsing header or footer
  class HeaderFooterChild < OOXMLDocumentObject
    # @return [Symbol] type of header
    attr_reader :type
    # @return [String] raw text of header
    attr_reader :raw_string

    def initialize(type: nil, parent: nil)
      @type = type
      super(parent: parent)
    end

    # Parse HeaderFooterChild data
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [HeaderFooterChild] result of parsing
    def parse(node)
      @raw_string = node.text
      self
    end
  end
end
