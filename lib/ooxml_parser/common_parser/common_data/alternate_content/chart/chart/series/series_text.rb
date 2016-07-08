require_relative 'series_text/string_reference'
module OoxmlParser
  # Class for parsing `c:tx` object
  class SeriesText < OOXMLDocumentObject
    # @return [StringReference] String reference of series
    attr_accessor :string

    # Parse Order
    # @param [Nokogiri::XML:Node] node with Order
    # @return [Order] result of parsing
    def self.parse(node)
      text = SeriesText.new
      node.xpath('*').each do |series_child|
        case series_child.name
        when 'strRef'
          text.string = StringReference.parse(series_child)
        end
      end
      text
    end
  end

  Categories = SeriesText
end
