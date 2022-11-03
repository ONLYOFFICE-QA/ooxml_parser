# frozen_string_literal: true

module OoxmlParser
  # Class for working with AlphaChannel
  class ColorAlphaChannel < OOXMLDocumentObject
    # @return [Integer] value of alpha channel
    attr_reader :value

    # Parse AlphaChannel value
    # @param node [Nokogiri::XML::Element] node to parse
    # @return [ColorAlphaChannel] parsed object
    def parse(node)
      alpha_channel_node = node.xpath('a:alpha', 'xmlns:a' => 'http://schemas.openxmlformats.org/drawingml/2006/main').first
      alpha_channel_node = node.xpath('w14:alpha', 'xmlns:w14' => 'http://schemas.microsoft.com/office/word/2010/wordml').first if alpha_channel_node.nil?
      alpha_channel_node.attributes.each do |key, value|
        case key
        when 'val'
          @value = (value.value.to_f / 1_000.0).round(0)
        end
      end
      @value ||= 100.0

      self
    end
  end
end
