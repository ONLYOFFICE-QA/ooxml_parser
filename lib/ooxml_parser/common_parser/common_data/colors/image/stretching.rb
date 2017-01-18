require_relative 'stretching/fill_rectangle'
module OoxmlParser
  # Class for parsing `stretch` tag
  class Stretching < OOXMLDocumentObject
    attr_accessor :fill_rectangle

    # Parse Stretching object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Stretching] result of parsing
    def parse(node)
      @fill_rectangle = FillRectangle.new(parent: self).parse(node.xpath('a:fillRect').first) if node.xpath('a:fillRect').first
      self
    end
  end
end
