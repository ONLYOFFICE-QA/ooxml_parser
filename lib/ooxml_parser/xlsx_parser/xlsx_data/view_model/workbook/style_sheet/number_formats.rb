require_relative 'number_formats/number_format'
module OoxmlParser
  # Parsing `numFmts` tag
  class NumberFormats < OOXMLDocumentObject
    # @return [Array, NumberFormat] array of number formats
    attr_accessor :number_formats_array

    def initialize(parent: nil)
      @number_formats_array = []
      @parent = parent
    end

    # @return [Array, NumberFormats] accessor
    def [](key)
      @number_formats_array[key]
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
        when 'numFmt'
          @number_formats_array << NumberFormat.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
