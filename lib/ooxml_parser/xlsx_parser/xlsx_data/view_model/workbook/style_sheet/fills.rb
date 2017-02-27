require_relative 'fills/fill'
module OoxmlParser
  # Parsing `fonts` tag
  class Fills < OOXMLDocumentObject
    # @return [Array, Fill] array of fills
    attr_accessor :fills_array

    def initialize(parent: nil)
      @fills_array = []
      @parent = parent
    end

    # @return [Array, Fill] accessor
    def [](key)
      @fills_array[key]
    end

    # Parse Fills data
    # @param [Nokogiri::XML:Element] node with Fills data
    # @return [Fills] value of Fills data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'count'
          @count = value.value.to_i
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'fill'
          @fills_array << Fill.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
