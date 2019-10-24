# frozen_string_literal: true

require_relative 'fonts/font'
module OoxmlParser
  # Parsing `fonts` tag
  class Fonts < OOXMLDocumentObject
    # @return [Array, Font] array of number formats
    attr_accessor :fonts_array

    def initialize(parent: nil)
      @fonts_array = []
      @parent = parent
    end

    # @return [Array, Font] accessor
    def [](key)
      @fonts_array[key]
    end

    # Parse NumberFormats data
    # @param [Nokogiri::XML:Element] node with NumberFormats data
    # @return [NumberFormats] value of NumberFormats data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'count'
          @count = value.value.to_i
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'font'
          @fonts_array << Font.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
