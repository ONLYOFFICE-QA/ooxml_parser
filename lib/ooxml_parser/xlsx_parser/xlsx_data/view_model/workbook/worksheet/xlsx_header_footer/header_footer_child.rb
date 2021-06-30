# frozen_string_literal: true

module OoxmlParser
  # Class for parsing header or footer
  class HeaderFooterChild < OOXMLDocumentObject
    # @return [Symbol] type of header
    attr_reader :type
    # @return [String] raw text of header
    attr_reader :raw_string
    # @return [String] left part of header
    attr_reader :left
    # @return [String] center part of header
    attr_reader :center
    # @return [String] right part of header
    attr_reader :right

    def initialize(type: nil, raw_string: nil, parent: nil)
      @type = type
      @raw_string = raw_string
      super(parent: parent)
    end

    # Parse HeaderFooterChild data
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [HeaderFooterChild] result of parsing
    def parse(node)
      @raw_string = node.text
      parse_raw_string
      self
    end

    # Perform parsing of header parts from string
    def parse_raw_string
      right = @raw_string.split('&R')
      @right = right.last
      center = right.first.split('&C')
      @center = center.last
      @left = center.first.split('&L').last
    end
  end
end
