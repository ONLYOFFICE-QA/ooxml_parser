# frozen_string_literal: true

module OoxmlParser
  # Class for parsing header or footer
  class HeaderFooterChild < OOXMLDocumentObject
    # @return [Symbol] type of header
    attr_reader :type
    # @return [String] raw text of header
    attr_reader :raw_string

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
      self
    end

    # @return [String] right part of header
    def right
      return @right if @right

      right = @raw_string.match(/&R(.+)/)
      return nil unless right

      @right = right[1]
    end

    # @return [String] center part of header
    def center
      return @center if @center

      center = @raw_string.split('&R').first.match(/&C(.+)/)
      return nil unless center

      @center = center[1]
    end

    def left
      return @left if @left

      left = @raw_string.gsub("&R#{right}", '')
      left = left.gsub("&C#{center}", '')
      return nil if left == ''

      left.gsub('&L', '')
    end
  end
end
