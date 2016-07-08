module OoxmlParser
  # Class for parsing `c:order` object
  class Order < OOXMLDocumentObject
    # @return [Integer] value of order
    attr_accessor :value

    # Parse Order
    # @param [Nokogiri::XML:Node] node with Order
    # @return [Order] result of parsing
    def self.parse(node)
      index = Order.new
      node.attributes.each do |key, value|
        case key
        when 'val'
          index.value = value.value.to_f
        end
      end
      index
    end
  end
end
